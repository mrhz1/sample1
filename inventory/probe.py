"""Pass 2 - open a sample of files to identify what they really are.

This is the only pass that opens files, so it is the expensive one on a
network share, and it is built around one observation: DICOM files sitting in
the same directory are almost always slices of the same study. So instead of
reading 3.5 million headers, it reads a handful per directory, checks that the
samples agree, and infers the rest.

When the samples disagree - mixed studies, or DICOM mixed with junk - that
directory is escalated and read in full. Rare, but real, and the alternative
is silently mis-counting it. Every study row records which happened, so you
can always tell an inferred count from a counted one.

Only files whose extension doesn't already identify them get opened. A .pdf
is a PDF; there is nothing to learn by reading it. In practice that means the
extensionless files, which is where the DICOMs live.

Resumable: each directory is committed as it finishes, so re-running picks up
where it stopped.

Usage:
    python probe.py [--db inventory.db] [--workers N] [--sample N] [--force]

    --db PATH     Database built by crawl.py (default: inventory.db).
    --workers N   Parallel file readers (default: 16; raise on a NAS).
    --sample N    Files to sample per directory (default: 4). Higher is safer
                  and slower; 4 catches mixed directories reliably because it
                  always includes the first and last file by name.
    --force       Re-probe directories already done.
"""

import argparse
import json
import os
import random
import sqlite3
import sys
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait

try:
    import pydicom
except ImportError:
    sys.exit("pydicom is required: pip install pydicom")

from crawl import long_path

BATCH_SECONDS = 5.0
PROGRESS_SECONDS = 15.0

# Extensions we trust without opening the file. Anything else gets sniffed.
KNOWN_EXTS = {
    "pdf": "pdf", "doc": "word", "docx": "word", "rtf": "word",
    "xls": "excel", "xlsx": "excel", "csv": "excel",
    "ppt": "slides", "pptx": "slides",
    "jpg": "image", "jpeg": "image", "png": "image", "gif": "image",
    "bmp": "image", "tif": "image", "tiff": "image",
    "zip": "archive", "rar": "archive", "7z": "archive", "gz": "archive",
    "tar": "archive", "tgz": "archive",
    "exe": "program", "msi": "program", "dll": "program", "bat": "program",
    "txt": "text", "log": "text", "ini": "text", "inf": "text", "xml": "text",
    "avi": "video", "mp4": "video", "mov": "video", "wmv": "video",
    "dcm": "dicom", "dicom": "dicom", "ima": "dicom",
}

# The matching only needs these three, but we keep the whole header anyway -
# see read_one(). Sampling means we store roughly one header per study, not
# one per slice, so the cost is small and the metadata is there when a later
# question needs a tag nobody thought about yet.
WANTED_TAGS = [(0x0008, 0x0020), (0x0008, 0x0060), (0x0020, 0x000D)]

SCHEMA = """
CREATE TABLE IF NOT EXISTS dir_probe (
    dir_id     INTEGER PRIMARY KEY,
    method     TEXT NOT NULL,     -- 'sampled', 'full', 'skipped'
    n_sampled  INTEGER NOT NULL,
    n_dicom    INTEGER NOT NULL,  -- inferred or counted DICOM files
    note       TEXT
);

-- Full header of every file we actually opened, as DICOM JSON. Pixel data is
-- never read. NOTE: this includes patient name, DOB and institution - the
-- database is PHI and should be stored accordingly.
CREATE TABLE IF NOT EXISTS dicom_meta (
    file_id INTEGER PRIMARY KEY,
    dir_id  INTEGER NOT NULL,
    json    TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS dicom_meta_dir ON dicom_meta(dir_id);

CREATE TABLE IF NOT EXISTS studies (
    id          INTEGER PRIMARY KEY,
    dir_id      INTEGER NOT NULL,
    study_uid   TEXT,
    study_date  TEXT,
    modality    TEXT,
    slice_count INTEGER NOT NULL,
    confidence  TEXT NOT NULL     -- 'counted' | 'inferred'
);
CREATE INDEX IF NOT EXISTS studies_dir ON studies(dir_id);
"""


def sniff_magic(head):
    """Identify a file from its first bytes. Returns a kind, or 'unknown'."""
    if len(head) >= 132 and head[128:132] == b"DICM":
        return "dicom"
    if head[:4] == b"%PDF":
        return "pdf"
    if head[:2] == b"PK":
        return "archive"          # also .docx/.xlsx; extension already caught those
    if head[:2] == b"MZ":
        return "program"
    if head[:4] == b"\xd0\xcf\x11\xe0":
        return "office"           # legacy .doc/.xls container
    if head[:3] == b"\xff\xd8\xff" or head[:8] == b"\x89PNG\r\n\x1a\n":
        return "image"
    if head[:6] in (b"GIF87a", b"GIF89a"):
        return "image"
    return "unknown"


def read_one(path, keep_full=True):
    """Return (kind, study_uid, study_date, modality, header_json).

    The whole header is read rather than just the three tags we match on.
    Reading the file is the expensive part; parsing the rest of the header
    once it is in memory is nearly free, and it means a later question about
    some other tag doesn't cost another pass over the share.
    """
    try:
        with open(long_path(path), "rb") as fh:
            head = fh.read(132)
            kind = sniff_magic(head)
            if kind != "dicom":
                return kind, None, None, None, None
            fh.seek(0)
            ds = pydicom.dcmread(fh, stop_before_pixels=True)
            blob = None
            if keep_full:
                try:
                    blob = json.dumps(ds.to_json_dict())
                except Exception:
                    # A malformed tag shouldn't cost us the study itself.
                    blob = None
            return (
                "dicom",
                str(getattr(ds, "StudyInstanceUID", "") or "") or None,
                str(getattr(ds, "StudyDate", "") or "") or None,
                str(getattr(ds, "Modality", "") or "") or None,
                blob,
            )
    except Exception as exc:
        return f"?error:{type(exc).__name__}", None, None, None, None


def pick_sample(names, k):
    """First and last by name, plus a spread of others.

    First and last matter: a directory holding two studies almost always has
    them ordered so that the boundary shows up between the extremes.
    """
    if len(names) <= k:
        return list(names)
    chosen = {0, len(names) - 1}
    rnd = random.Random(len(names))
    while len(chosen) < k:
        chosen.add(rnd.randrange(len(names)))
    return [names[i] for i in sorted(chosen)]


def probe_dir(dir_path, candidates, sample_size, mode="all"):
    """Probe one directory. `candidates` is [(file_id, name), ...].

    Returns (method, results, note) where results maps file_id -> read_one
    tuple.

    mode="all"     read every file and keep every header. Slower and much
                   bigger, but the metadata is then there for any question
                   asked later without touching the share again - and study
                   grouping becomes counted rather than inferred.
    mode="sample"  read a few files per directory, keep their headers, and
                   infer the rest of the directory from them.
    mode="none"    sample, and don't keep headers at all.
    """
    if mode == "all":
        results = {fid: read_one(os.path.join(dir_path, name), True)
                   for fid, name in candidates}
        return "full", results, None

    names = [n for _, n in candidates]
    by_name = {n: fid for fid, n in candidates}
    sample = pick_sample(sorted(names), sample_size)

    keep_full = mode == "sample"
    results = {}
    for name in sample:
        results[by_name[name]] = read_one(os.path.join(dir_path, name), keep_full)

    kinds = {r[0] for r in results.values()}
    uids = {r[1] for r in results.values() if r[0] == "dicom"}

    agreed = kinds == {"dicom"} and len(uids) == 1
    if agreed or len(candidates) <= sample_size:
        return ("sampled" if agreed else "full"), results, None

    # Samples disagree and there's more to read - do the whole directory.
    for fid, name in candidates:
        if fid not in results:
            results[fid] = read_one(os.path.join(dir_path, name))
    return "full", results, "samples disagreed; directory read in full"


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--workers", type=int, default=16)
    ap.add_argument("--sample", type=int, default=4)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--metadata", default="all",
                    choices=["all", "sample", "none"],
                    help="all: read every DICOM and keep every header (default; "
                         "biggest and slowest, but nothing needs re-reading "
                         "later). sample: a few per directory. none: identify "
                         "files only, keep no headers.")
    ap.add_argument("--probe", type=int, default=0,
                    help="stop after N directories and report throughput - "
                         "use this to time the share before the full run")
    args = ap.parse_args()

    conn = sqlite3.connect(args.db)
    conn.executescript(SCHEMA)
    if args.force:
        conn.executescript(
            "DELETE FROM dir_probe; DELETE FROM studies; "
            "DELETE FROM dicom_meta; "
            "UPDATE files SET kind = NULL, kind_source = NULL;")
        conn.commit()

    # Directories holding at least one file we can't identify by extension.
    unknown_exts = ",".join("?" * len(KNOWN_EXTS))
    todo = conn.execute(
        f"""SELECT d.id, d.path
              FROM dirs d
             WHERE d.scanned = 1
               AND EXISTS (SELECT 1 FROM files f
                            WHERE f.dir_id = d.id
                              AND f.ext NOT IN ({unknown_exts}))
               AND d.id NOT IN (SELECT dir_id FROM dir_probe)
          ORDER BY d.id""",
        list(KNOWN_EXTS),
    ).fetchall()

    done_already = conn.execute("SELECT COUNT(*) FROM dir_probe").fetchone()[0]
    print(f"{len(todo):,} directories to probe"
          + (f" ({done_already:,} already done)" if done_already else ""))
    if not todo:
        return

    started = time.time()
    last_commit = last_report = time.time()
    n_done = n_files_read = 0

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        in_flight = {}
        queue = list(reversed(todo))
        while queue or in_flight:
            while queue and len(in_flight) < args.workers * 2:
                dir_id, dir_path = queue.pop()
                candidates = conn.execute(
                    f"SELECT id, name FROM files "
                    f"WHERE dir_id = ? AND ext NOT IN ({unknown_exts})",
                    [dir_id] + list(KNOWN_EXTS),
                ).fetchall()
                fut = pool.submit(probe_dir, dir_path, candidates,
                                  args.sample, args.metadata)
                in_flight[fut] = (dir_id, candidates)

            done, _ = wait(list(in_flight), return_when=FIRST_COMPLETED)
            for fut in done:
                dir_id, candidates = in_flight.pop(fut)
                method, results, note = fut.result()
                n_files_read += len(results)

                for fid, r in results.items():
                    if r[0] == "dicom" and r[4]:
                        conn.execute(
                            "INSERT OR REPLACE INTO dicom_meta VALUES (?,?,?)",
                            (fid, dir_id, r[4]))

                dicom_reads = [r for r in results.values() if r[0] == "dicom"]
                if method == "sampled" and dicom_reads:
                    # Samples agreed: the whole directory is that one study.
                    _, uid, date, modality, _ = dicom_reads[0]
                    n_dicom = len(candidates)
                    conn.execute(
                        "INSERT INTO studies(dir_id, study_uid, study_date, "
                        "modality, slice_count, confidence) VALUES (?,?,?,?,?,?)",
                        (dir_id, uid, date, modality, n_dicom, "inferred"))
                    conn.executemany(
                        "UPDATE files SET kind = ?, kind_source = ? WHERE id = ?",
                        [("dicom", "inferred", fid) for fid, _ in candidates])
                else:
                    # Read in full: group the real results into studies.
                    n_dicom = len(dicom_reads)
                    groups = {}
                    for fid, (kind, uid, date, modality, _blob) in results.items():
                        conn.execute(
                            "UPDATE files SET kind = ?, kind_source = ? "
                            "WHERE id = ?", (kind, "magic", fid))
                        if kind == "dicom":
                            groups.setdefault((uid, date, modality), 0)
                            groups[(uid, date, modality)] += 1
                    for (uid, date, modality), count in groups.items():
                        conn.execute(
                            "INSERT INTO studies(dir_id, study_uid, study_date,"
                            " modality, slice_count, confidence) "
                            "VALUES (?,?,?,?,?,?)",
                            (dir_id, uid, date, modality, count, "counted"))

                conn.execute(
                    "INSERT OR REPLACE INTO dir_probe VALUES (?,?,?,?,?)",
                    (dir_id, method, len(results), n_dicom, note))
                n_done += 1

            now = time.time()
            if now - last_commit >= BATCH_SECONDS:
                conn.commit()
                last_commit = now
            if args.probe and n_done >= args.probe:
                conn.commit()
                el = max(now - started, 0.001)
                print(f"\n-- probe stopped after {n_done:,} directories --")
                print(f"  {n_files_read:,} files opened in {el:.0f}s "
                      f"({n_files_read / el:,.0f} files/s)")
                size = conn.execute(
                    "SELECT COALESCE(SUM(LENGTH(json)),0) FROM dicom_meta"
                ).fetchone()[0]
                if n_files_read:
                    print(f"  metadata so far: {size / 1e6:,.1f} MB "
                          f"({size / max(n_files_read,1):,.0f} bytes/file)")
                print("\nRe-run without --probe to continue; it resumes here.")
                return
            if now - last_report >= PROGRESS_SECONDS:
                el = max(now - started, 0.001)
                pct = 100.0 * n_done / len(todo)
                eta = (len(todo) - n_done) / (n_done / el) / 3600 if n_done else 0
                print(f"  {n_done:,}/{len(todo):,} dirs ({pct:.1f}%)  "
                      f"{n_files_read:,} files opened  eta~{eta:.1f}h", flush=True)
                last_report = now

    conn.commit()
    # Built now rather than in the schema: maintaining it through millions of
    # UPDATEs above would cost more than building it once at the end.
    conn.execute("CREATE INDEX IF NOT EXISTS files_kind ON files(kind)")
    conn.commit()
    n_studies, n_slices = conn.execute(
        "SELECT COUNT(*), COALESCE(SUM(slice_count),0) FROM studies").fetchone()
    escalated = conn.execute(
        "SELECT COUNT(*) FROM dir_probe WHERE method = 'full'").fetchone()[0]
    print(f"\ndone in {(time.time() - started) / 60:.1f} min")
    print(f"  {n_done:,} directories probed, {n_files_read:,} files actually opened")
    print(f"  {n_studies:,} studies, {n_slices:,} DICOM files")
    if args.metadata == "all":
        print("  every DICOM read - study counts are exact, and any tag can be "
              "pulled later\n  with extract.py without touching the archive again")
    n_meta = conn.execute("SELECT COUNT(*) FROM dicom_meta").fetchone()[0]
    if args.metadata != "all":
        print(f"  {escalated:,} directories needed a full read (samples disagreed)")
    meta_bytes = conn.execute(
        "SELECT COALESCE(SUM(LENGTH(json)),0) FROM dicom_meta").fetchone()[0]
    print(f"  {n_meta:,} full DICOM headers stored "
          f"({meta_bytes / 1e6:,.1f} MB of JSON in dicom_meta)")
    if args.metadata != "all" and n_files_read < n_slices:
        print(f"  sampling saved ~{n_slices - n_files_read:,} file opens")
    conn.close()


if __name__ == "__main__":
    main()
