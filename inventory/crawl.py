"""Pass 1 - walk an archive and record every file in a SQLite database.

This is the "dumb" pass: it opens no files and interprets nothing. It only
enumerates directories and records name/size/mtime, which on Windows comes
back with the directory listing for free (no extra round trip per file).
That matters on a network share, where per-file latency - not bandwidth - is
what makes a multi-million-file walk slow.

Everything that requires actually opening a file (magic-byte sniffing, DICOM
headers) happens in a later pass, driven by queries against this database, so
an interrupted run never loses the walk.

Resumable: a directory is marked scanned in the same transaction that inserts
its contents, so a crash, a dropped share or a sleeping laptop resumes from the
last committed batch rather than starting over. Just run the same command again.

Usage:
    python crawl.py <root> [--db inventory.db] [--workers N] [--expect-files N]

    root              Folder to walk (local path, mapped drive or UNC share).
    --db PATH         Database to build (default: inventory.db).
    --workers N       Parallel directory readers (default: 16). Raise to 32-64
                      on a NAS, where latency dominates; drop to 4 on a local
                      disk, where it doesn't.
    --expect-files N  Rough total file count, used only to turn the running
                      throughput into an ETA.
    --probe N         Stop after N directories and report throughput. Use this
                      first on a slow share to find out what you're in for.

Examples:
    python crawl.py "\\\\nas\\studies" --db inventory.db --workers 32 --probe 2000
    python crawl.py "\\\\nas\\studies" --db inventory.db --workers 32 --expect-files 5000000
    python crawl.py "D:\\Archive" --db inventory.db
"""

import argparse
import os
import sqlite3
import sys
import time
from collections import deque
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait

# Commit roughly this often. Bigger batches mean fewer fsyncs and a faster
# walk; smaller batches mean less rescanning after a crash.
BATCH_DIRS = 200
BATCH_SECONDS = 5.0
PROGRESS_SECONDS = 15.0

SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA synchronous  = NORMAL;

CREATE TABLE IF NOT EXISTS dirs (
    id        INTEGER PRIMARY KEY,
    path      TEXT UNIQUE NOT NULL,
    parent_id INTEGER,
    depth     INTEGER NOT NULL,
    scanned   INTEGER NOT NULL DEFAULT 0,
    error     TEXT,
    code      TEXT              -- patient code, filled by report.py
);
CREATE INDEX IF NOT EXISTS dirs_pending ON dirs(scanned) WHERE scanned = 0;

-- One row per file. crawl.py fills the first six columns; probe.py fills
-- kind/kind_source; report.py fills the code columns. They are kept on this
-- one table so "SELECT ... FROM files WHERE code = 'AA0001'" needs no join.
CREATE TABLE IF NOT EXISTS files (
    id     INTEGER PRIMARY KEY,
    dir_id INTEGER NOT NULL,
    name   TEXT NOT NULL,
    ext    TEXT NOT NULL,
    size   INTEGER,
    mtime  REAL,

    kind        TEXT,           -- dicom, pdf, word, ...      (probe.py)
    kind_source TEXT,           -- 'ext' | 'magic' | 'inferred'

    code          TEXT,         -- display form, e.g. AA0001  (report.py)
    prefix        TEXT,         -- AA
    number        INTEGER,      -- 1  <- the real identity
    code_source   TEXT,         -- 'filename' | 'folder' | 'none'
    conflict_with TEXT          -- other code seen, when name and folder differ
);
CREATE INDEX IF NOT EXISTS files_dir ON files(dir_id);
CREATE INDEX IF NOT EXISTS files_ext ON files(ext);
-- files_kind and files_code are created by probe.py and report.py once those
-- columns hold data; building them here would only slow the bulk insert.

CREATE TABLE IF NOT EXISTS meta (
    key   TEXT PRIMARY KEY,
    value TEXT
);
"""


def long_path(path):
    """Windows caps paths at 260 chars unless they carry the \\\\?\\ prefix.

    A messy archive will exceed that, and the failure is silent - the
    directory just appears empty - so every path handed to the OS goes
    through here. Paths are stored in the database unprefixed.
    """
    if os.name != "nt":
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def scan_one(path):
    """Read one directory. Returns (subdir names, file tuples, error)."""
    subdirs, files = [], []
    try:
        with os.scandir(long_path(path)) as it:
            for entry in it:
                try:
                    # follow_symlinks=False keeps us out of junction loops and,
                    # on Windows, uses the stat data the listing already
                    # returned instead of making another round trip.
                    if entry.is_dir(follow_symlinks=False):
                        subdirs.append(entry.name)
                    else:
                        st = entry.stat(follow_symlinks=False)
                        ext = os.path.splitext(entry.name)[1].lower().lstrip(".")
                        files.append((entry.name, ext, st.st_size, st.st_mtime))
                except OSError as exc:
                    # One unreadable entry shouldn't lose the whole directory.
                    files.append((entry.name, "?error", None, None))
                    del exc
    except OSError as exc:
        return [], [], f"{type(exc).__name__}: {exc}"
    return subdirs, files, None


def open_db(db_path):
    conn = sqlite3.connect(db_path)
    conn.executescript(SCHEMA)
    return conn


def ensure_dir(conn, path, parent_id, depth):
    """Insert a directory if new, and return its id either way.

    Idempotent because a resumed run re-scans any directory whose transaction
    didn't commit, and will rediscover children that did.
    """
    conn.execute(
        "INSERT INTO dirs(path, parent_id, depth) VALUES (?,?,?) "
        "ON CONFLICT(path) DO NOTHING",
        (path, parent_id, depth),
    )
    row = conn.execute("SELECT id FROM dirs WHERE path = ?", (path,)).fetchone()
    return row[0]


def crawl(root, db_path, workers, expect_files, probe):
    root = os.path.abspath(root)
    conn = open_db(db_path)
    conn.execute(
        "INSERT INTO meta(key, value) VALUES ('root', ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (root,),
    )
    ensure_dir(conn, root, None, 0)
    conn.commit()

    pending = conn.execute(
        "SELECT id, path, depth FROM dirs WHERE scanned = 0 ORDER BY id"
    ).fetchall()
    already = conn.execute("SELECT COUNT(*) FROM dirs WHERE scanned = 1").fetchone()[0]
    n_files = conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
    if already:
        print(f"resuming: {already:,} directories already scanned, "
              f"{n_files:,} files recorded, {len(pending):,} directories queued")

    queue = deque(pending)
    started = time.time()
    dirs_done = 0
    files_done = 0
    last_commit = time.time()
    last_report = time.time()
    uncommitted = 0

    with ThreadPoolExecutor(max_workers=workers) as pool:
        in_flight = {}
        while queue or in_flight:
            # Keep the pool fed. Two tasks per worker is enough to hide
            # latency without queueing up a huge amount of unwritten work.
            while queue and len(in_flight) < workers * 2:
                d = queue.popleft()
                in_flight[pool.submit(scan_one, d[1])] = d

            done, _ = wait(list(in_flight), return_when=FIRST_COMPLETED)
            for fut in done:
                dir_id, dir_path, depth = in_flight.pop(fut)
                subdirs, files, error = fut.result()

                # Delete-then-insert so re-scanning an uncommitted directory
                # can't double-count its files.
                conn.execute("DELETE FROM files WHERE dir_id = ?", (dir_id,))
                if files:
                    conn.executemany(
                        "INSERT INTO files(dir_id, name, ext, size, mtime) "
                        "VALUES (?,?,?,?,?)",
                        [(dir_id, *f) for f in files],
                    )
                new_children = []
                for name in subdirs:
                    child_path = os.path.join(dir_path, name)
                    child_id = ensure_dir(conn, child_path, dir_id, depth + 1)
                    new_children.append((child_id, child_path, depth + 1))
                conn.execute(
                    "UPDATE dirs SET scanned = 1, error = ? WHERE id = ?",
                    (error, dir_id),
                )

                queue.extend(new_children)
                dirs_done += 1
                files_done += len(files)
                uncommitted += 1

            now = time.time()
            if uncommitted >= BATCH_DIRS or now - last_commit >= BATCH_SECONDS:
                conn.commit()
                uncommitted = 0
                last_commit = now

            if now - last_report >= PROGRESS_SECONDS:
                report(started, dirs_done, files_done, len(queue), expect_files)
                last_report = now

            if probe and dirs_done >= probe:
                conn.commit()
                print(f"\n-- probe stopped after {dirs_done:,} directories --")
                report(started, dirs_done, files_done, len(queue), expect_files)
                print("\nRun the same command without --probe to continue; the "
                      "walk resumes from here.")
                return

    conn.commit()
    total_dirs, total_files = conn.execute(
        "SELECT (SELECT COUNT(*) FROM dirs), (SELECT COUNT(*) FROM files)"
    ).fetchone()
    errors = conn.execute(
        "SELECT COUNT(*) FROM dirs WHERE error IS NOT NULL"
    ).fetchone()[0]
    elapsed = time.time() - started
    print(f"\ndone in {elapsed / 60:.1f} min")
    print(f"  {total_dirs:,} directories, {total_files:,} files")
    if errors:
        print(f"  {errors:,} directories could not be read (see the dirs.error "
              f"column - usually permissions or a path that's still too long)")
    conn.close()


def report(started, dirs_done, files_done, queued, expect_files):
    elapsed = max(time.time() - started, 0.001)
    dps = dirs_done / elapsed
    fps = files_done / elapsed
    line = (f"  {dirs_done:,} dirs / {files_done:,} files  "
            f"({dps:,.0f} dirs/s, {fps:,.0f} files/s)  queue={queued:,}")
    if expect_files and fps > 0:
        remaining = max(expect_files - files_done, 0) / fps
        line += f"  eta~{remaining / 3600:.1f}h"
    print(line, flush=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("root")
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--workers", type=int, default=16)
    ap.add_argument("--expect-files", type=int, default=0)
    ap.add_argument("--probe", type=int, default=0)
    args = ap.parse_args()

    if not os.path.isdir(args.root):
        sys.exit(f"not a directory: {args.root}")
    crawl(args.root, args.db, args.workers, args.expect_files, args.probe)


if __name__ == "__main__":
    main()
