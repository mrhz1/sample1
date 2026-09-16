"""Compare the stored DICOM headers of two files and list only what differs.

Written for one question: are these two files the same image filed twice, or
two different images that happen to look alike? A full header is hundreds of
tags, almost all identical between copies, so printing everything buries the
answer. This prints the differences and nothing else.

Reads the headers probe.py stored - it never opens the files, so it works on
an archive that is offline or slow.

Usage:
    python compare_headers.py --db inventory.db 12345 67890
    python compare_headers.py --db inventory.db 12345 67890 --all

    --db PATH   Database, after crawl.py and probe.py --metadata all.
    --all       Also list the tags that are identical.
    --binary    Compare binary elements too (pixel data is never stored, but
                overlays and private blobs can be). Off by default: they are
                large and a difference in them is rarely what you are after.

Finding the two ids - every copy of one image, with its file_id:

    SELECT u.uid, f.id, d.path || '/' || f.name
      FROM dicom_uid u JOIN files f ON f.id = u.file_id
      JOIN dirs d ON d.id = f.dir_id
     WHERE u.uid IN (SELECT uid FROM dicom_uid GROUP BY uid HAVING COUNT(*) > 1)
     ORDER BY u.uid;
"""

import argparse
import json
import os
import sqlite3
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from extract import describe, value_of  # noqa: E402  - reuse, don't reimplement

# Tags that SHOULD differ between two copies of one image, and tags that must
# not. Knowing which is which is the whole point of the comparison, so the
# output says so rather than leaving it to the reader.
IDENTITY_TAGS = {
    "00080018": "SOPInstanceUID - same value means the SAME image",
    "0020000D": "StudyInstanceUID",
    "0020000E": "SeriesInstanceUID",
    "00200013": "InstanceNumber",
    "00080020": "StudyDate",
    "00080060": "Modality",
    "00100020": "PatientID",
    "00100010": "PatientName",
}


def load(conn, file_id):
    row = conn.execute(
        "SELECT d.path || '/' || f.name, f.size, m.json"
        "  FROM files f JOIN dirs d ON d.id = f.dir_id"
        "  LEFT JOIN dicom_meta m ON m.file_id = f.id"
        " WHERE f.id = ?", (file_id,)).fetchone()
    if row is None:
        sys.exit(f"no file with id {file_id} in this database")
    path, size, blob = row
    if blob is None:
        sys.exit(f"file {file_id} has no stored header ({path}).\n"
                 "Only files probe.py opened have one - re-run it with "
                 "--metadata all to store every header.")
    return path, size, json.loads(blob)


def is_binary(header, tag):
    item = header.get(tag) or {}
    return "InlineBinary" in item and "Value" not in item


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("file_ids", nargs=2, type=int)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--binary", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    conn = sqlite3.connect(f"file:{os.path.abspath(args.db)}?mode=ro", uri=True)
    a_id, b_id = args.file_ids
    a_path, a_size, a = load(conn, a_id)
    b_path, b_size, b = load(conn, b_id)
    conn.close()

    print(f"A  file_id {a_id:<10} {a_size:>12,} bytes  {a_path}")
    print(f"B  file_id {b_id:<10} {b_size:>12,} bytes  {b_path}")
    print()

    same, diffs, only_a, only_b = [], [], [], []
    for tag in sorted(set(a) | set(b)):
        if not args.binary and (is_binary(a, tag) or is_binary(b, tag)):
            continue
        in_a, in_b = tag in a, tag in b
        va, vb = value_of(a, tag), value_of(b, tag)
        if in_a and not in_b:
            only_a.append((tag, va))
        elif in_b and not in_a:
            only_b.append((tag, vb))
        elif va == vb:
            same.append((tag, va))
        else:
            diffs.append((tag, va, vb))

    def show(tag, va, vb=None, side="A"):
        note = IDENTITY_TAGS.get(tag)
        print(f"  {tag}  {describe(tag)[:34]:<34}")
        print(f"      {side}: {str(va)[:96]}")
        if vb is not None:
            print(f"      B: {str(vb)[:96]}")
        if note:
            print(f"      ^ {note}")

    if diffs:
        print(f"DIFFERENT VALUES ({len(diffs)})")
        print("-" * 78)
        for tag, va, vb in diffs:
            show(tag, va, vb)
    if only_a:
        print(f"\nPRESENT ONLY IN A ({len(only_a)})")
        print("-" * 78)
        for tag, va in only_a:
            show(tag, va)
    if only_b:
        print(f"\nPRESENT ONLY IN B ({len(only_b)})")
        print("-" * 78)
        for tag, vb in only_b:
            show(tag, vb, side="B")

    if not diffs and not only_a and not only_b:
        print("NO DIFFERENCES - the two stored headers are identical"
              + (" (binary elements not compared; pass --binary)"
                 if not args.binary else ""))

    if args.all and same:
        print(f"\nIDENTICAL ({len(same)})")
        print("-" * 78)
        for tag, va in same:
            print(f"  {tag}  {describe(tag)[:34]:<34} {str(va)[:40]}")

    # The verdict people actually came for.
    sop_a, sop_b = value_of(a, "00080018"), value_of(b, "00080018")
    print("\n" + "=" * 78)
    if sop_a and sop_a == sop_b:
        print("SAME IMAGE - identical SOPInstanceUID. One of these is a copy.")
        if a_size != b_size:
            print(f"  Note: the files differ in size ({a_size:,} vs {b_size:,}"
                  " bytes) - same image, re-encoded or with tags added.")
    elif sop_a and sop_b:
        print("DIFFERENT IMAGES - the SOPInstanceUIDs do not match.")
    else:
        print("CANNOT TELL - at least one header has no SOPInstanceUID.")
        print("  Compare the tags above, and treat equal name+size as evidence,"
              " not proof.")
    print(f"  {len(same):,} tags identical, {len(diffs):,} different, "
          f"{len(only_a) + len(only_b):,} present on one side only")


if __name__ == "__main__":
    main()
