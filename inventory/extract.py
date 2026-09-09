"""Pull any DICOM field for every image in the archive, straight from the DB.

This is the payoff for storing full headers in pass 2: when someone asks
"give me the manufacturer / slice thickness / study description for every
image", it is a local query taking seconds, not another multi-hour crawl of
the share.

Usage:
    python extract.py --db inventory.db --list-tags
    python extract.py --db inventory.db --tags PatientName,StudyDescription
    python extract.py --db inventory.db --tags 00181030,SliceThickness \\
                      --out fields.csv --by study

    --list-tags   Show every tag present in the stored headers, with how many
                  files carry it. Run this first to see what you actually have.
    --tags LIST   Comma-separated DICOM keywords (StudyDescription) or hex tags
                  (00081030). Mixed freely.
    --by file     One row per DICOM file (default).
    --by study    One row per study - far fewer rows, and right for anything
                  that doesn't vary slice to slice.
    --code CODE   Restrict to one patient.
    --out PATH    CSV to write (default: stdout).
"""

import argparse
import csv
import json
import sqlite3
import sys
import time
from collections import Counter

try:
    from pydicom.datadict import dictionary_description, tag_for_keyword
except ImportError:
    sys.exit("pydicom is required: pip install pydicom")


def resolve(name):
    """Accept either a keyword (StudyDescription) or a hex tag (00081030)."""
    clean = name.strip().replace(",", "").replace("(", "").replace(")", "")
    if len(clean) == 8:
        try:
            int(clean, 16)
            return clean.upper()
        except ValueError:
            pass
    tag = tag_for_keyword(name.strip())
    if tag is None:
        sys.exit(f"unknown DICOM keyword or tag: {name}")
    return f"{tag:08X}"


def describe(hex_tag):
    try:
        return dictionary_description(int(hex_tag, 16))
    except Exception:
        return "?"


def list_tags(conn):
    counts, examples = Counter(), {}
    for (blob,) in conn.execute("SELECT json FROM dicom_meta"):
        try:
            header = json.loads(blob)
        except Exception:
            continue
        for tag, item in header.items():
            counts[tag] += 1
            if tag not in examples:
                value = item.get("Value") or [""]
                examples[tag] = str(value[0])[:28]
    print(f"{'TAG':<10} {'FILES':>9}  {'DESCRIPTION':<34} EXAMPLE")
    print("-" * 90)
    for tag, n in counts.most_common():
        print(f"{tag:<10} {n:>9,}  {describe(tag)[:34]:<34} {examples.get(tag,'')}")
    print(f"\n{len(counts)} distinct tags across "
          f"{conn.execute('SELECT COUNT(*) FROM dicom_meta').fetchone()[0]:,} headers")


def value_of(header, hex_tag):
    item = header.get(hex_tag)
    if not item:
        return ""
    value = item.get("Value")
    if not value:
        return ""
    if len(value) == 1:
        v = value[0]
        # PersonName comes back as {"Alphabetic": "..."}.
        return v.get("Alphabetic", "") if isinstance(v, dict) else str(v)
    return "\\".join(str(v) for v in value)


def show_one(conn, which):
    """Print one stored header as tag / VR / description / value."""
    row = conn.execute(
        "SELECT m.file_id, f.name, d.path, m.json FROM dicom_meta m "
        "JOIN files f ON f.id = m.file_id JOIN dirs d ON d.id = f.dir_id "
        "WHERE m.file_id = ? OR f.name = ? LIMIT 1",
        (which if str(which).isdigit() else -1, which)).fetchone()
    if not row:
        print(f"no stored header for {which!r}", file=sys.stderr)
        return
    fid, name, path, blob = row
    header = json.loads(blob)
    print(f"# file_id {fid}  {path}/{name}")
    print(f"# {len(header)} tags stored\n")
    for tag in sorted(header):
        v = value_of(header, tag)
        vr = header[tag].get("vr", "")
        pretty = f"({tag[:4]},{tag[4:]})"
        print(f"{pretty}  {vr:2}  {describe(tag)[:38]:<38} {str(v)[:60]}")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--list-tags", action="store_true")
    ap.add_argument("--tags")
    ap.add_argument("--by", default="file", choices=["file", "study"])
    ap.add_argument("--code")
    ap.add_argument("--out")
    ap.add_argument("--show", metavar="FILE_ID_OR_NAME",
                    help="print one stored header in readable form, the way "
                         "pydicom shows it, instead of writing a CSV")
    args = ap.parse_args()

    conn = sqlite3.connect(args.db)
    if args.list_tags:
        list_tags(conn)
        return
    if args.show:
        show_one(conn, args.show)
        return
    if not args.tags:
        sys.exit("give --tags, or --list-tags to see what's available")

    tags = [resolve(t) for t in args.tags.split(",") if t.strip()]
    labels = [t.strip() for t in args.tags.split(",") if t.strip()]

    sql = """SELECT m.file_id, f.code, d.path, f.name, m.json
               FROM dicom_meta m
               JOIN files f ON f.id = m.file_id
               JOIN dirs  d ON d.id = f.dir_id"""
    params = []
    if args.code:
        sql += " WHERE f.code = ?"
        params.append(args.code)

    out = open(args.out, "w", newline="", encoding="utf-8") if args.out else sys.stdout
    writer = csv.writer(out)
    writer.writerow(["code", "folder", "file"] + labels)

    n_total = conn.execute(
        "SELECT COUNT(*) FROM dicom_meta").fetchone()[0]
    seen = set()
    n = 0
    n_seen = 0
    t_started = t_last = time.time()
    for _fid, code, path, name, blob in conn.execute(sql, params):
        n_seen += 1
        if n_seen % 50000 == 0:
            now = time.time()
            if now - t_last >= 5.0:
                rate = n_seen / max(now - t_started, 0.001)
                eta = (n_total - n_seen) / rate / 60 if rate else 0
                print(f"  {n_seen:,}/{n_total:,} headers "
                      f"({100.0 * n_seen / max(n_total,1):.1f}%)  "
                      f"eta~{eta:.1f} min", file=sys.stderr, flush=True)
                t_last = now
        try:
            header = json.loads(blob)
        except Exception:
            continue
        values = [value_of(header, t) for t in tags]
        if args.by == "study":
            # One row per study: the folder plus the study UID identifies it,
            # and everything else in that folder is the same study.
            key = (path, value_of(header, "0020000D"))
            if key in seen:
                continue
            seen.add(key)
            name = ""
        writer.writerow([code or "", path, name] + values)
        n += 1

    if args.out:
        out.close()
        print(f"wrote {args.out}  ({n:,} rows)")


if __name__ == "__main__":
    main()
