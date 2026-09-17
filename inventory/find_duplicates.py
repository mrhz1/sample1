"""Find DICOM files that are copies of each other, without touching the share.

The same study often lands in the archive more than once - a working copy, a
re-export, a folder duplicated during a migration. Counting files then says
30,000 images where a person would say 14,000.

`probe.py --metadata all` stored every DICOM header, so this is answerable
locally. Two files with the same **SOPInstanceUID** are the same image: it is
generated once by the scanner, mandatory in every conformant file, and
preserved by a copy. Comparing bytes would be equally certain and would mean
re-reading a terabyte; this reads the database instead.

Where a header has no SOPInstanceUID - some anonymisers strip it - the fallback
is SeriesInstanceUID plus InstanceNumber, which is unique within a series for
the same reason. A file with neither is given an identity unique to itself and
so is never claimed as a copy of anything; the count of those is reported, so
"no duplicates" is never confused with "could not tell".

File name and size are deliberately NOT used. Slice names repeat across studies
(every series has an IM00001) and slices of one modality are often identical in
size, so that rule merges unrelated images and undercounts the archive.

Results are cached in a `dicom_uid` table so patient_summary.py can show the
per-patient column without redoing the work.

Usage:
    python find_duplicates.py --db inventory.db
    python find_duplicates.py --db inventory.db --out duplicates.xlsx

    --db PATH     Database, after crawl.py, probe.py and report.py.
    --out PATH    Also write the duplicate groups to Excel.
    --rebuild     Recompute the cached identities from scratch.
    --limit N     Groups to write to Excel (default 5000, newest first).
    --match MODE  How alike two files must be to count as duplicates:
                    image (default) - the same image, by SOPInstanceUID or
                      series + instance number. A copy that was re-encoded or
                      had a tag edited still counts: it is the same picture.
                    exact - the same image AND the same file name AND the same
                      byte size AND a byte-identical stored header. Answers
                      "is one of these safe to delete" rather than "is this
                      the same image", and will not flag an edited copy.
    --same-code   Only count two files as duplicates when both are attributed
                  to the same patient. Without it a copy filed under another
                  code still counts toward the archive total, and is reported
                  separately as a filing error.
    --explain A B Say why one pair of file_ids is, or is not, counted as a
                  duplicate. Use it when two files look identical but the
                  per-patient count does not move.
"""

import argparse
import os
import sqlite3
import sys
import time

SCHEMA = """
CREATE TABLE IF NOT EXISTS dicom_uid (
    file_id INTEGER PRIMARY KEY,
    uid     TEXT NOT NULL,   -- SOPInstanceUID, or a series+instance key, or
                             -- 'file:<id>' when the header identifies nothing
    source  TEXT NOT NULL    -- 'sop' | 'series+instance' | 'none'
);
CREATE INDEX IF NOT EXISTS dicom_uid_uid ON dicom_uid(uid);
"""


def build_identities(conn, rebuild):
    """One identity per DICOM file, cached. Returns (n_sop, n_fallback)."""
    conn.executescript(SCHEMA)
    if rebuild:
        conn.execute("DELETE FROM dicom_uid")
    done = conn.execute("SELECT COUNT(*) FROM dicom_uid").fetchone()[0]
    total = conn.execute(
        "SELECT COUNT(*) FROM files WHERE kind = 'dicom'").fetchone()[0]
    if done and done >= total:
        print(f"  using {done:,} cached identities (--rebuild to redo)")
    else:
        started = time.time()
        # json_extract over millions of headers is the slow part, so it runs
        # once and the answer is kept.
        # 'file:<id>' is unique per row on purpose: a file whose header
        # identifies nothing can then never collide with another, so it is
        # counted as its own image rather than silently merged into one.
        conn.execute("""
            INSERT OR REPLACE INTO dicom_uid(file_id, uid, source)
            SELECT f.id,
                   CASE
                     WHEN json_extract(m.json, '$."00080018".Value[0]') IS NOT NULL
                       THEN json_extract(m.json, '$."00080018".Value[0]')
                     WHEN json_extract(m.json, '$."0020000E".Value[0]') IS NOT NULL
                      AND json_extract(m.json, '$."00200013".Value[0]') IS NOT NULL
                       THEN 'series:'
                            || json_extract(m.json, '$."0020000E".Value[0]')
                            || '#'
                            || json_extract(m.json, '$."00200013".Value[0]')
                     ELSE 'file:' || f.id
                   END,
                   CASE
                     WHEN json_extract(m.json, '$."00080018".Value[0]') IS NOT NULL
                       THEN 'sop'
                     WHEN json_extract(m.json, '$."0020000E".Value[0]') IS NOT NULL
                      AND json_extract(m.json, '$."00200013".Value[0]') IS NOT NULL
                       THEN 'series+instance'
                     ELSE 'none'
                   END
              FROM files f
              LEFT JOIN dicom_meta m ON m.file_id = f.id
             WHERE f.kind = 'dicom'
        """)
        conn.commit()
        print(f"  identified {total:,} DICOM files in "
              f"{time.time() - started:.1f}s")
    return dict(conn.execute(
        "SELECT source, COUNT(*) FROM dicom_uid GROUP BY source"))


def refine_exact(conn):
    """Split each candidate group by name, size and the stored header.

    Only files that already share an image identity are compared, so this is a
    short second pass rather than a scan of the archive. Two copies that
    differ in any of the three stop being duplicates under --match exact.
    """
    groups = conn.execute("""
        SELECT uid FROM dicom_uid GROUP BY uid HAVING COUNT(*) > 1
    """).fetchall()
    changed = 0
    for (uid,) in groups:
        rows = conn.execute("""
            SELECT u.file_id, f.name, f.size, COALESCE(m.json, '')
              FROM dicom_uid u JOIN files f ON f.id = u.file_id
              LEFT JOIN dicom_meta m ON m.file_id = u.file_id
             WHERE u.uid = ?""", (uid,)).fetchall()
        buckets = {}
        for file_id, name, size, blob in rows:
            buckets.setdefault((name, size, blob), []).append(file_id)
        if len(buckets) == 1:
            continue
        # The largest bucket keeps the original identity, so the answer does
        # not depend on the order rows came back in.
        order = sorted(buckets.values(), key=len, reverse=True)
        for n, ids in enumerate(order[1:], start=1):
            conn.executemany("UPDATE dicom_uid SET uid = ? WHERE file_id = ?",
                             [(f"{uid}#{n}", fid) for fid in ids])
        changed += 1
    conn.commit()
    return changed


def summarise(conn, same_code=False):
    """Totals, and the per-patient count of redundant copies."""
    if same_code:
        # Count an image once per patient, so a copy filed under a different
        # code is not counted as redundant against this one.
        total, unique = conn.execute("""
            SELECT COUNT(*), COUNT(DISTINCT f.code || '|' || u.uid)
              FROM dicom_uid u JOIN files f ON f.id = u.file_id
             WHERE f.code IS NOT NULL""").fetchone()
    else:
        total, unique = conn.execute(
            "SELECT COUNT(*), COUNT(DISTINCT uid) FROM dicom_uid").fetchone()

    # Copies of one image filed under one patient: the extra ones are waste.
    per_patient = dict(conn.execute("""
        SELECT code, SUM(n - 1) FROM (
            SELECT f.code AS code, u.uid AS uid, COUNT(*) AS n
              FROM dicom_uid u JOIN files f ON f.id = u.file_id
             WHERE f.code IS NOT NULL
             GROUP BY f.code, u.uid HAVING COUNT(*) > 1)
         GROUP BY code
    """))

    # The same image filed under two different patients is not waste, it is a
    # filing error, and only a person can say which side is wrong.
    cross = conn.execute("""
        SELECT COUNT(*) FROM (
            SELECT u.uid FROM dicom_uid u JOIN files f ON f.id = u.file_id
             WHERE f.code IS NOT NULL
             GROUP BY u.uid HAVING COUNT(DISTINCT f.code) > 1)
    """).fetchone()[0]
    return total, unique, per_patient, cross


def groups(conn, limit):
    """Duplicate groups, biggest first, for the Excel output."""
    return conn.execute("""
        SELECT u.uid,
               COUNT(*) AS copies,
               COUNT(DISTINCT f.code) AS patients,
               GROUP_CONCAT(DISTINCT COALESCE(f.code, '(unassigned)')),
               MIN(u.source),
               GROUP_CONCAT(d.path || '/' || f.name, '  |  ')
          FROM dicom_uid u
          JOIN files f ON f.id = u.file_id
          JOIN dirs d ON d.id = f.dir_id
         GROUP BY u.uid HAVING COUNT(*) > 1
         ORDER BY copies DESC, u.uid
         LIMIT ?
    """, (limit,)).fetchall()


def write_workbook(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = "Duplicate Groups"
    header = ["Image ID", "Copies", "Patients", "Patient Codes",
              "Identified By", "Files"]
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for uid, copies, patients, codes, source, files in rows:
        ws.append([uid, copies, patients, codes,
                   {"sop": "SOPInstanceUID",
                    "series+instance": "series + instance number"}.get(
                        source, "not identifiable"),
                   files])
    for letter, width in zip("ABCDEF", (46, 8, 9, 22, 18, 120)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ws.max_row}"
    wb.save(path)


def explain(conn, a_id, b_id):
    """Why a pair is or is not counted. Every reason it can fail, checked."""
    # The likeliest reason a pair is not counted is that this table was never
    # built, so create it empty rather than failing on a missing table - the
    # verdict below then says exactly that.
    conn.executescript(SCHEMA)
    rows = {}
    for fid in (a_id, b_id):
        row = conn.execute(
            "SELECT f.id, f.code, f.kind, u.uid, u.source, d.path || '/' || f.name"
            "  FROM files f JOIN dirs d ON d.id = f.dir_id"
            "  LEFT JOIN dicom_uid u ON u.file_id = f.id"
            " WHERE f.id = ?", (fid,)).fetchone()
        if row is None:
            print(f"  file_id {fid} is not in this database")
            return
        rows[fid] = row
    for fid in (a_id, b_id):
        _, code, kind, uid, source, path = rows[fid]
        print(f"  file_id {fid}")
        print(f"    patient code : {code or '(none - unassigned)'}")
        print(f"    kind         : {kind or '(not probed)'}")
        print(f"    image id     : {uid or '(none - run without --explain first)'}")
        print(f"    identified by: {source or '-'}")
        print(f"    path         : {path}")
    a, b = rows[a_id], rows[b_id]
    print("\n  verdict")
    if a[3] is None or b[3] is None:
        print("    NOT COUNTED - no cached identity. Run find_duplicates.py")
        print("    without --explain first; it builds the dicom_uid table.")
        return
    if a[2] != "dicom" or b[2] != "dicom":
        print("    NOT COUNTED - only files probe.py identified as DICOM are")
        print(f"    compared, and these are {a[2]!r} and {b[2]!r}.")
        return
    if "none" in (a[4], b[4]):
        print("    NOT COUNTED - at least one header identifies nothing: no")
        print("    SOPInstanceUID and no series + instance number. Such a file")
        print("    is counted as its own image rather than guessed at.")
        return
    if a[3] != b[3]:
        print("    NOT COUNTED - different images. The two headers carry")
        print("    different identifiers, so these are not copies.")
        return
    if a[1] != b[1]:
        print("    SAME IMAGE, but filed under two DIFFERENT patient codes:")
        print(f"        {a[1] or '(unassigned)'}   vs   {b[1] or '(unassigned)'}")
        print("    Counted as a cross-patient duplicate, NOT in either")
        print("    patient's 'Duplicate DICOM Images' - that column counts")
        print("    redundant copies within one patient, and this pair is a")
        print("    filing error instead: one of the two codes is wrong.")
        print("    Look at the two paths above and see which folder misleads.")
        return
    print(f"    COUNTED - same image, same patient ({a[1]}). One is redundant.")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--out", default=None)
    ap.add_argument("--rebuild", action="store_true")
    ap.add_argument("--limit", type=int, default=5000)
    ap.add_argument("--explain", nargs=2, type=int, metavar=("A", "B"))
    ap.add_argument("--match", default="image", choices=["image", "exact"])
    ap.add_argument("--same-code", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    conn = sqlite3.connect(args.db)
    have = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table'")}
    if "dicom_meta" not in have:
        sys.exit("no dicom_meta table - run probe.py first")

    if args.explain:
        explain(conn, *args.explain)
        conn.close()
        return

    # The cache holds identities built under one mode; switching modes without
    # rebuilding would report the previous mode's answer.
    previous = conn.execute(
        "SELECT value FROM meta WHERE key = 'duplicate_match'").fetchone()
    rebuild = args.rebuild or (previous and previous[0] != args.match)
    if previous and previous[0] != args.match and not args.rebuild:
        print(f"  --match changed ({previous[0]} -> {args.match}), rebuilding")
    sources = build_identities(conn, rebuild)
    if args.match == "exact":
        split = refine_exact(conn)
        print(f"  --match exact: {split:,} group(s) split by name, size or"
              " header - same image, different file")
    conn.execute("INSERT OR REPLACE INTO meta(key, value) VALUES"
                 " ('duplicate_match', ?)", (args.match,))
    conn.commit()
    total, unique, per_patient, cross = summarise(conn, args.same_code)
    redundant = total - unique

    scope = "same patient only" if args.same_code else "whole archive"
    print(f"\n  match: {args.match}   scope: {scope}")
    print(f"\n  {total:,} DICOM files")
    print(f"  {unique:,} distinct images")
    print(f"  {redundant:,} redundant copies "
          f"({100.0 * redundant / total:.1f}% of the files)"
          if total else "  nothing to compare")
    print(f"\n  identified by SOPInstanceUID:        {sources.get('sop', 0):,}")
    if sources.get("series+instance"):
        print(f"  identified by series + instance no: "
              f"{sources['series+instance']:,}")
    if sources.get("none"):
        print(f"  NOT identifiable from the header:   {sources['none']:,}")
        print("    counted as distinct images - they cannot be checked, which"
              " is not the\n    same as having no duplicates. Re-probe with"
              " --metadata all if these\n    files were probed with"
              " --metadata sample or none.")
    if cross:
        print(f"\n  {cross:,} image(s) filed under MORE THAN ONE patient - a"
              " filing error, not waste")
    worst = sorted(per_patient.items(), key=lambda kv: -kv[1])[:5]
    if worst:
        print("\n  most affected patients:")
        for code, n in worst:
            print(f"    {code:<12} {n:,} redundant copies")

    if args.out:
        write_workbook(groups(conn, args.limit), args.out)
        print(f"\n  duplicate groups: {args.out}")
    print("\n  cached in dicom_uid - patient_summary.py will show the"
          " per-patient column now")
    conn.close()


if __name__ == "__main__":
    main()
