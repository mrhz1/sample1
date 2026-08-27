"""Merge every per-code results workbook into one master Excel file.

Reads all `<code>-results.xlsx` files produced by match_reports.py and writes
a single workbook with two sheets:

  "All Data" - every row from every file, same columns and same format.
  "Totals"   - plain counts: matched vs not matched, and a per-modality
               breakdown.

Usage:
    python merge_results.py [results_dir] [output.xlsx]

    results_dir    Folder holding the *-results.xlsx files (default: ./results)
    output.xlsx    Master workbook to write (default: <results_dir>/../master_report.xlsx)

Examples:
    python merge_results.py
    python merge_results.py results master_report.xlsx
    python merge_results.py "/mnt/data/results" "/mnt/data/master_report.xlsx"
"""

import os
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

# Must stay in step with match_reports.py's SUMMARY_HEADER / SUMMARY_WIDTHS.
SUMMARY_HEADER = [
    "Code",
    "Study Date",
    "Modality",
    "DICOM File Count",
    "Matched PDF File(s)",
    "Match Status",
]
SUMMARY_WIDTHS = (12, 14, 10, 16, 45, 40)
# The stats block at the bottom of a sheet needs a roomier first column than
# the Code column does, so the data sheet gets a widened column A.
ALL_DATA_WIDTHS = (26,) + SUMMARY_WIDTHS[1:]
TOTALS_WIDTHS = (52, 14, 12, 14, 16, 19, 20)

MATCHED_STATUS = "Matched"
NO_DICOM_STATUS = "No DICOM images found"

BOLD = Font(bold=True)


def read_result_rows(path):
    """Read back one per-code results file, minus its header row."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        return [list(row) for row in ws.iter_rows(min_row=2, values_only=True)]
    finally:
        wb.close()


def to_int(value):
    """DICOM File Count is blank on PDF-only rows; treat that as 0."""
    if value in (None, ""):
        return 0
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def collect_rows(results_dir):
    """Load every per-code workbook in the folder, sorted by filename."""
    rows = []
    files = sorted(results_dir.glob("*-results.xlsx"))
    if not files:
        print(f"No *-results.xlsx files found in {results_dir}")
        return rows, 0
    for path in files:
        try:
            file_rows = read_result_rows(path)
        except Exception as exc:
            print(f"  WARNING: could not read {path.name}: {exc}")
            continue
        print(f"  {path.name}: {len(file_rows):,} row(s)")
        rows.extend(file_rows)
    return rows, len(files)


def tally(rows):
    """Boil the merged rows down to the numbers for the Totals sheet."""
    totals = {
        "rows": len(rows),
        "codes": len({r[0] for r in rows if r[0]}),
        "dicom_total": 0,
        "dicom_matched": 0,
        "dicom_unmatched": 0,
        "studies_matched": 0,
        "studies_unmatched": 0,
        "pdf_only_rows": 0,
    }
    statuses = Counter()
    per_modality = defaultdict(
        lambda: {
            "rows": 0,
            "studies": 0,
            "dicom": 0,
            "dicom_matched": 0,
            "dicom_unmatched": 0,
            "pdf_only": 0,
        }
    )

    for row in rows:
        modality = (row[2] or "(blank)").strip() or "(blank)"
        count = to_int(row[3])
        status = (row[5] or "").strip()
        statuses[status or "(blank)"] += 1

        bucket = per_modality[modality]
        bucket["rows"] += 1
        bucket["dicom"] += count
        totals["dicom_total"] += count

        if status == NO_DICOM_STATUS:
            totals["pdf_only_rows"] += 1
            bucket["pdf_only"] += 1
            continue

        bucket["studies"] += 1
        if status == MATCHED_STATUS:
            totals["studies_matched"] += 1
            totals["dicom_matched"] += count
            bucket["dicom_matched"] += count
        else:
            totals["studies_unmatched"] += 1
            totals["dicom_unmatched"] += count
            bucket["dicom_unmatched"] += count

    return totals, statuses, per_modality


def stat_lines(totals, file_count):
    """The overall numbers, in one place so the terminal and the Excel
    summary block can never drift apart."""
    return [
        ("files merged", file_count),
        ("codes", totals["codes"]),
        ("rows", totals["rows"]),
        ("DICOM files total", totals["dicom_total"]),
        ("DICOM files matched", totals["dicom_matched"]),
        ("DICOM files not matched", totals["dicom_unmatched"]),
        ("studies matched", totals["studies_matched"]),
        ("studies not matched", totals["studies_unmatched"]),
        ("PDFs with no DICOM", totals["pdf_only_rows"]),
    ]


MODALITY_COLUMNS = [
    "Modality",
    "Found (rows)",
    "Studies",
    "DICOM files",
    "DICOM matched",
    "DICOM not matched",
    "PDF only (no DICOM)",
]


def modality_lines(per_modality):
    """Per-modality totals across every code - Echo, CT, MRI and the rest."""
    out = []
    for modality in sorted(per_modality):
        b = per_modality[modality]
        out.append(
            [
                modality,
                b["rows"],
                b["studies"],
                b["dicom"],
                b["dicom_matched"],
                b["dicom_unmatched"],
                b["pdf_only"],
            ]
        )
    return out


def append_stats_block(ws, totals, statuses, per_modality, file_count):
    """Drop the same numbers the terminal prints at the bottom of a sheet."""

    def heading(text):
        ws.append([])
        ws.append([text])
        ws.cell(row=ws.max_row, column=1).font = BOLD

    def header_row(labels):
        ws.append(labels)
        for cell in ws[ws.max_row]:
            cell.font = BOLD

    heading("SUMMARY")
    for label, value in stat_lines(totals, file_count):
        ws.append([label, value])

    heading("BY MATCH STATUS")
    header_row(["Status", "Rows"])
    for status, count in statuses.most_common():
        ws.append([status, count])

    heading("BY MODALITY (all codes)")
    header_row(MODALITY_COLUMNS)
    for line in modality_lines(per_modality):
        ws.append(line)


def write_master(path, rows, file_count):
    """Write the master workbook via a temp file, then swap it in."""
    totals, statuses, per_modality = tally(rows)

    wb = Workbook()
    ws = wb.active
    ws.title = "All Data"
    ws.append(SUMMARY_HEADER)
    for cell in ws[1]:
        cell.font = BOLD
    for row in rows:
        ws.append(row)
    for column, width in zip("ABCDEFG", ALL_DATA_WIDTHS):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"

    # Same numbers, twice: at the bottom of the data sheet (scroll to the
    # end and they're right there) and on their own sheet.
    append_stats_block(ws, totals, statuses, per_modality, file_count)

    ts = wb.create_sheet("Totals")
    for column, width in zip("ABCDEFG", TOTALS_WIDTHS):
        ts.column_dimensions[column].width = width
    append_stats_block(ts, totals, statuses, per_modality, file_count)
    ts.delete_rows(1)  # the block opens with a spacer row; not needed here

    tmp = path.with_suffix(path.suffix + ".tmp")
    wb.save(tmp)
    os.replace(tmp, path)
    return totals, per_modality


def main(argv):
    if argv and argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    results_dir = Path(argv[0]) if argv else Path.cwd() / "results"
    if not results_dir.is_dir():
        print(f"Results folder not found: {results_dir}")
        return 1

    if len(argv) > 1:
        output_xlsx = Path(argv[1])
    else:
        output_xlsx = results_dir.parent / "master_report.xlsx"
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    print(f"Reading: {results_dir}")
    rows, file_count = collect_rows(results_dir)
    if not rows:
        print("Nothing to merge.")
        return 1

    totals, per_modality = write_master(output_xlsx, rows, file_count)

    print("")
    print(f"Master file: {output_xlsx}")
    for label, value in stat_lines(totals, file_count):
        print(f"  {label:<24}{value:,}")
    print("")
    print("  By modality (all codes):")
    width = max([len(m) for m in per_modality] + [8])
    print(
        f"    {'modality':<{width}}  {'found':>7}{'studies':>9}{'dicom':>9}"
        f"{'matched':>9}{'unmatched':>11}{'pdf only':>10}"
    )
    for line in modality_lines(per_modality):
        modality, found, studies, dicom, matched, unmatched, pdf_only = line
        print(
            f"    {modality:<{width}}  {found:>7,}{studies:>9,}{dicom:>9,}"
            f"{matched:>9,}{unmatched:>11,}{pdf_only:>10,}"
        )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
