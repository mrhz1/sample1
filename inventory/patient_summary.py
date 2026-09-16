"""A two-sheet workbook: the headline numbers, and one line per patient.

Answers "what exactly does AA0001 have?" in a single row rather than across
several sheets:

    AA0001   600 DICOM images, 7 PDFs
             400 DICOMs have a report, 200 do not
             3 reports have no DICOM | 610 files, 1.2 GB

The counts here are **DICOM files**, not studies. master_report.xlsx counts
studies, because that is the unit a report is written about; this counts the
files underneath them, because that is the unit the archive is measured in. A
patient with one reported study of 400 slices and one unreported study of 200
appears here as 400 and 200, and there as 1 and 1. Both are right; this is the
one people mean when they ask how much is outstanding.

Reads the matching report.py already wrote to the database, so the two files
can never disagree and nothing is re-matched here.

Usage:
    python report.py --db inventory.db --prefixes AA   # must run first
    python patient_summary.py --db inventory.db --out patient_summary.xlsx

    --db PATH        Database, after crawl.py, probe.py and report.py.
    --out PATH       Workbook to write (default patient_summary.xlsx).
    --prefix LIST    Only these prefixes, comma-separated.
    --find-duplicates
                     Build the duplicate index first if it is missing. Off by
                     default because it is a one-off pass over every stored
                     header - minutes on a large archive - and this script is
                     otherwise instant. Without it the duplicate column reads
                     "not checked" rather than a misleading 0.
"""

import argparse
import os
import re
import sqlite3
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEAD_FILL = PatternFill("solid", fgColor="DDDDDD")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")

# status values written by report.py
S_BOTH = "report and image"
S_IMAGE = "only image"
S_REPORT = "only report"
S_OTHER = "other pdf"

# Spelled out, no abbreviations - these headings get read by people who do not
# work with the archive daily, and "DICOMs w/o Report" means nothing to them.
HEADER = [
    ("Patient Code", 14),
    ("Total DICOM Images", 17),
    ("Duplicate DICOM Images", 18),
    ("Total PDFs", 12),
    ("DICOMs With a Report", 18),
    ("DICOMs With No Report", 18),
    ("Reports With No DICOM", 18),
    ("Other PDFs", 12),
    ("Total Files (All Types)", 18),
    ("Total Size", 13),
]


def human_bytes(n):
    size = float(n or 0)
    for unit in ("B", "KB", "MB", "GB", "TB", "PB"):
        if size < 1024 or unit == "PB":
            return f"{size:,.1f} {unit}"
        size /= 1024


def code_prefix(code):
    m = re.match(r"[A-Za-z]+", code or "")
    return m.group(0).upper() if m else ""


NOT_CHECKED = "not checked"


def collect(conn):
    """(code -> every number on the patient line, whether duplicates are known)."""
    rows = {}

    def row(code):
        return rows.setdefault(code, {
            "dicom_reported": 0, "reports_covering": set(),
            "dicom_unreported": 0, "reports_orphan": 0, "other_pdf": 0,
            "other_files": 0, "dicom_files": 0, "pdf_files": 0,
            "duplicates": 0, "total_files": 0, "size": 0})

    for code, status, slices, report in conn.execute(
            "SELECT code, status, slices, report FROM study_report"
            " WHERE code IS NOT NULL"):
        r = row(code)
        if status == S_BOTH:
            r["dicom_reported"] += slices or 0
            if report:
                r["reports_covering"].add(report)
        elif status == S_IMAGE:
            r["dicom_unreported"] += slices or 0
        elif status == S_REPORT:
            r["reports_orphan"] += 1
        elif status == S_OTHER:
            r["other_pdf"] += 1

    # Redundant copies, if the duplicate index exists. Said to be unknown
    # rather than reported as zero when it does not - a zero answers the
    # question wrongly, and it is the answer people act on.
    have = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table'")}
    dupes_known = "dicom_uid" in have and conn.execute(
        "SELECT 1 FROM dicom_uid LIMIT 1").fetchone() is not None
    if dupes_known:
        for code, n in conn.execute("""
                SELECT code, SUM(n - 1) FROM (
                    SELECT f.code AS code, u.uid AS uid, COUNT(*) AS n
                      FROM dicom_uid u JOIN files f ON f.id = u.file_id
                     WHERE f.code IS NOT NULL
                     GROUP BY f.code, u.uid HAVING COUNT(*) > 1)
                 GROUP BY code"""):
            row(code)["duplicates"] = n or 0

    for code, kind, n, size in conn.execute(
            "SELECT code, kind, COUNT(*), COALESCE(SUM(size), 0) FROM files"
            " WHERE code IS NOT NULL GROUP BY code, kind"):
        r = row(code)
        r["total_files"] += n
        r["size"] += size
        if kind == "dicom":
            r["dicom_files"] += n
        elif kind == "pdf":
            r["pdf_files"] += n
        else:
            r["other_files"] += n
    return rows, dupes_known


def build_rows(rows, dupes_known=True):
    """Totals first, then the split - so a reader can see the whole before
    the parts, and spot it when the parts do not add up to it."""
    out = []
    for code in sorted(rows):
        r = rows[code]
        out.append([code, r["dicom_files"],
                    r["duplicates"] if dupes_known else NOT_CHECKED,
                    r["pdf_files"],
                    r["dicom_reported"], r["dicom_unreported"],
                    r["reports_orphan"], r["other_pdf"],
                    r["total_files"], human_bytes(r["size"])])
    return out


def overview(conn, rows, root, dupes_known=True):
    total = lambda k: sum(r[k] for r in rows.values())  # noqa: E731
    reported = total("dicom_reported")
    unreported = total("dicom_unreported")
    both = sum(1 for r in rows.values()
               if (r["dicom_reported"] or r["dicom_unreported"])
               and (r["reports_covering"] or r["reports_orphan"]))
    only_img = sum(1 for r in rows.values()
                   if (r["dicom_reported"] or r["dicom_unreported"])
                   and not (r["reports_covering"] or r["reports_orphan"]))
    only_rep = sum(1 for r in rows.values()
                   if not (r["dicom_reported"] or r["dicom_unreported"])
                   and (r["reports_covering"] or r["reports_orphan"]))
    neither = len(rows) - both - only_img - only_rep
    pct = f"{100.0 * reported / (reported + unreported):.1f}%" if (
        reported + unreported) else "-"
    return [
        ("Archive root", root),
        ("Patients", len(rows)),
        ("", ""),
        ("Image files (DICOM) that have a report", reported),
        ("Image files with no report", unreported),
        ("Share of image files reported", pct),
        ("Duplicate image files (redundant copies)",
         sum(r["duplicates"] for r in rows.values()) if dupes_known
         else "not checked - run find_duplicates.py"),
        ("Image files not in any study", sum(
            max(r["dicom_files"] - r["dicom_reported"] - r["dicom_unreported"], 0)
            for r in rows.values())),
        ("", ""),
        ("Reports that cover images", sum(
            len(r["reports_covering"]) for r in rows.values())),
        ("Reports with no images", total("reports_orphan")),
        ("Other PDFs (no date - probably not reports)", total("other_pdf")),
        ("Other files (not images, not PDFs)", total("other_files")),
        ("", ""),
        ("Patients with both images and reports", both),
        ("Patients with images but no reports", only_img),
        ("Patients with reports but no images", only_rep),
        ("Patients with neither images nor reports", neither),
        ("", ""),
        ("Total files", total("total_files")),
        ("Total size", human_bytes(total("size"))),
    ]


def sheet(wb, title, header, widths, data, freeze=True):
    ws = wb.create_sheet(title)
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in data:
        ws.append(row)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    if freeze:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{ws.max_row}"
    return ws


def attribution_banner(conn):
    """When the codes in this database were written, and by what command.

    Printed because these tools display report.py's attribution rather than
    deriving their own - so a stale database looks exactly like a bug in the
    tool that reads it.
    """
    meta = dict(conn.execute(
        "SELECT key, value FROM meta WHERE key IN"
        " ('report_run_at', 'report_args')"))
    when = meta.get("report_run_at")
    if not when:
        return ("codes were written by a report.py run older than this stamp -"
                " re-run it if anything below looks stale")
    return f"codes written by report.py at {when}  ({meta.get('report_args', '')})"


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--out", default="patient_summary.xlsx")
    ap.add_argument("--prefix", default=None)
    ap.add_argument("--find-duplicates", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    # Read-only by default: this script only reports. --find-duplicates is the
    # one thing it can be asked to compute, and that needs to write its cache.
    if args.find_duplicates:
        import find_duplicates
        rw = sqlite3.connect(args.db)
        if "dicom_meta" not in {r[0] for r in rw.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'")}:
            sys.exit("no dicom_meta table - run probe.py --metadata all first")
        print("building the duplicate index (one-off)...")
        find_duplicates.build_identities(rw, rebuild=False)
        rw.close()

    conn = sqlite3.connect(f"file:{os.path.abspath(args.db)}?mode=ro", uri=True)
    have = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table'")}
    if "study_report" not in have:
        sys.exit("no study_report table - run report.py --prefixes ... first")

    root = conn.execute("SELECT value FROM meta WHERE key='root'").fetchone()[0]
    banner = attribution_banner(conn)
    rows, dupes_known = collect(conn)
    if args.prefix:
        wanted = {p.strip().upper() for p in args.prefix.split(",") if p.strip()}
        rows = {c: r for c, r in rows.items() if code_prefix(c) in wanted}
        if not rows:
            print(f"warning: no codes with prefix {args.prefix}", file=sys.stderr)
    conn.close()

    wb = Workbook()
    wb.remove(wb.active)
    ws = sheet(wb, "Overview", ["Measure", "Value"], (44, 60),
               overview(None, rows, root, dupes_known), freeze=False)
    ws.auto_filter.ref = None

    data = build_rows(rows, dupes_known)
    ws = sheet(wb, "Patients", [h for h, _ in HEADER], [w for _, w in HEADER],
               data)
    # Anything outstanding on a row is worth the eye landing on it.
    for row in ws.iter_rows(min_row=2):
        if (row[5].value or 0) or (row[6].value or 0):
            for cell in row:
                cell.fill = WARN_FILL
    wb.save(args.out)

    print(f"wrote {args.out}")
    print(f"  {banner}")
    print(f"  {len(data):,} patients")
    print(f"  {sum(r[4] for r in data):,} DICOMs have a report, "
          f"{sum(r[5] for r in data):,} do not")
    print(f"  {sum(r[6] for r in data):,} reports with no DICOM, "
          f"{sum(r[7] for r in data):,} other PDFs")
    if not dupes_known:
        print("  duplicate images: NOT CHECKED - the column says so rather than"
              " showing 0.\n"
              "    python find_duplicates.py --db <db>        (once, then re-run"
              " this)\n"
              "    or re-run this with --find-duplicates")
    else:
        dupes = sum(r[2] for r in data if isinstance(r[2], int))
        print(f"  {dupes:,} redundant DICOM copies")


if __name__ == "__main__":
    main()
