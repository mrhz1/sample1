"""Find DICOM files that are copies of each other, without touching the share.

The same study often lands in the archive more than once - a working copy, a
re-export, a folder duplicated during a migration. Counting files then says
30,000 images where a person would say 14,000.

`probe.py --metadata all` stored every DICOM header, so this is answerable
locally. Two files with the same **SOPInstanceUID** are the same image: it is
generated once by the scanner, mandatory in every conformant file, and
preserved by a copy. Comparing bytes would be equally certain and would mean
re-reading a terabyte; this reads the database instead.

Where a header has no SOPInstanceUID - non-conformant or anonymised writers -
it falls back to file name plus exact byte size, which is good evidence but not
proof. Every count is reported split by which method produced it, so a fallback
number is never mistaken for a certain one.

Results are cached in a `dicom_uid` table so patient_summary.py can show the
per-patient column without redoing the work.

Usage:
    python find_duplicates.py --db inventory.db
    python find_duplicates.py --db inventory.db --out duplicates.xlsx

    --db PATH     Database, after crawl.py, probe.py and report.py.
    --out PATH    Also write the duplicate groups to Excel.
    --rebuild     Recompute the cached identities from scratch.
    --limit N     Groups to write to Excel (default 5000, newest first).
"""

import argparse
import os
import sqlite3
import sys
import time

SCHEMA = """
CREATE TABLE IF NOT EXISTS dicom_uid (
    file_id INTEGER PRIMARY KEY,
    uid     TEXT NOT NULL,   -- SOPInstanceUID, or 'name|size' when absent
    source  TEXT NOT NULL    -- 'sop' | 'name+size'
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
        conn.execute("""
            INSERT OR REPLACE INTO dicom_uid(file_id, uid, source)
            SELECT f.id,
                   COALESCE(json_extract(m.json, '$."00080018".Value[0]'),
                            f.name || '|' || COALESCE(f.size, -1)),
                   CASE WHEN json_extract(m.json, '$."00080018".Value[0]')
                        IS NOT NULL THEN 'sop' ELSE 'name+size' END
              FROM files f
              LEFT JOIN dicom_meta m ON m.file_id = f.id
             WHERE f.kind = 'dicom'
        """)
        conn.commit()
        print(f"  identified {total:,} DICOM files in "
              f"{time.time() - started:.1f}s")
    return dict(conn.execute(
        "SELECT source, COUNT(*) FROM dicom_uid GROUP BY source"))


def summarise(conn):
    """Totals, and the per-patient count of redundant copies."""
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
                   "SOPInstanceUID" if source == "sop" else "file name + size",
                   files])
    for letter, width in zip("ABCDEF", (46, 8, 9, 22, 18, 120)):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:F{ws.max_row}"
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--out", default=None)
    ap.add_argument("--rebuild", action="store_true")
    ap.add_argument("--limit", type=int, default=5000)
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    conn = sqlite3.connect(args.db)
    have = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table'")}
    if "dicom_meta" not in have:
        sys.exit("no dicom_meta table - run probe.py first")

    sources = build_identities(conn, args.rebuild)
    total, unique, per_patient, cross = summarise(conn)
    redundant = total - unique

    print(f"\n  {total:,} DICOM files")
    print(f"  {unique:,} distinct images")
    print(f"  {redundant:,} redundant copies "
          f"({100.0 * redundant / total:.1f}% of the files)"
          if total else "  nothing to compare")
    print(f"\n  identified by SOPInstanceUID: {sources.get('sop', 0):,}")
    if sources.get("name+size"):
        print(f"  identified by file name + size: {sources['name+size']:,}"
              "  <- evidence, not proof")
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
