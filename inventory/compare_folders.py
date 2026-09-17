"""How many files in one folder are copies of files in another?

    python compare_folders.py --db inventory.db \\
        --path1 "/nas/AA0976 base" --path2 "/nas/AA0976 backup" --code AA0976

Answers one question with one number: of the DICOM files under path2, how many
are copies of a file under path1. Both folders are read from the database, so
the share is never touched.

Two files are the same image when they carry the same SOPInstanceUID - an
identifier the scanner generates once per image, mandatory in a conformant
file, and preserved by copying. Where that is absent the fallback is
SeriesInstanceUID + InstanceNumber, unique within a series. A file with
neither is never claimed as a copy of anything, and is counted separately so
that "could not tell" is not read as "not a duplicate".

File names and sizes are deliberately not compared: every series has an
IM00001 and slices of one modality are often identical in size, so that rule
merges unrelated images.

    --db PATH     Database, after crawl.py, probe.py and find_duplicates.py.
    --path1 P     The folder the originals are expected in.
    --path2 P     The folder to check against it.
    --code C      Only files attributed to this patient code. Optional; without
                  it every file under the two folders is compared.
    --flat        Only the folders themselves, not their subfolders.
    --list N      Also print N example duplicate file names.
"""

import argparse
import os
import sqlite3
import sys


def normalise(path):
    return path.replace("\\", "/").rstrip("/")


def files_under(conn, path, code, flat):
    """{identity -> [file names]} for the DICOM files under one folder."""
    path = normalise(path)
    where = "d.path = ?" if flat else "(d.path = ? OR d.path LIKE ? || '/%')"
    params = [path] if flat else [path, path]
    if code:
        where += " AND f.code = ?"
        params.append(code)
    rows = conn.execute(
        f"""SELECT u.uid, u.source, f.name
              FROM files f JOIN dirs d ON d.id = f.dir_id
              LEFT JOIN dicom_uid u ON u.file_id = f.id
             WHERE {where} AND f.kind = 'dicom'""", params).fetchall()
    found, unknown = {}, 0
    for uid, source, name in rows:
        if uid is None or source == "none":
            unknown += 1
            continue
        found.setdefault(uid, []).append(name)
    return found, unknown, len(rows)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--path1", required=True)
    ap.add_argument("--path2", required=True)
    ap.add_argument("--code", default=None)
    ap.add_argument("--flat", action="store_true")
    ap.add_argument("--list", type=int, default=0)
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    conn = sqlite3.connect(f"file:{os.path.abspath(args.db)}?mode=ro", uri=True)
    if "dicom_uid" not in {r[0] for r in conn.execute(
            "SELECT name FROM sqlite_master WHERE type = 'table'")}:
        sys.exit("no dicom_uid table - run find_duplicates.py --db "
                 f"{args.db} first; it builds the image index this reads.")

    a, a_unknown, a_files = files_under(conn, args.path1, args.code, args.flat)
    b, b_unknown, b_files = files_under(conn, args.path2, args.code, args.flat)
    conn.close()

    if not a_files:
        sys.exit(f"no DICOM files found under path1 - check the path is exactly"
                 f" as the database stores it, and the --code:\n  {args.path1}")
    if not b_files:
        sys.exit(f"no DICOM files found under path2:\n  {args.path2}")

    shared = set(a) & set(b)
    # One image can sit in a folder more than once, so count files, not images.
    duplicate_files = sum(len(b[uid]) for uid in shared)

    print(f"\n  DUPLICATED FILES: {duplicate_files:,}")
    print(f"\n  of the {b_files:,} DICOM files under path2, {duplicate_files:,}"
          " are copies of an image")
    print(f"  that also exists under path1"
          + (f" for patient {args.code}" if args.code else ""))
    print(f"\n  path1  {a_files:>9,} files  {len(a):>9,} distinct images"
          f"   {args.path1}")
    print(f"  path2  {b_files:>9,} files  {len(b):>9,} distinct images"
          f"   {args.path2}")
    print(f"  shared {len(shared):>9,} images in both")
    print(f"  only in path2 {len(set(b) - set(a)):,} images - what path2 adds")
    if a_unknown or b_unknown:
        print(f"\n  NOT CHECKED: {a_unknown + b_unknown:,} file(s) whose header"
              " identifies nothing.\n  They are excluded from the count above"
              " rather than guessed at.")
    if args.list and shared:
        print(f"\n  examples:")
        for uid in list(shared)[:args.list]:
            print(f"    {b[uid][0]}  (also in path1 as {a[uid][0]})")


if __name__ == "__main__":
    main()
