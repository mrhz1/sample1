"""
Match PDF reports to DICOM studies and record the exact file-level linkage.

Expects two folder trees sharing the same subfolder naming (e.g. "R123"):
    images/R123/...   -> DICOM files (mixed in with other junk, e.g. viewer apps)
    reports/R123/...  -> PDF report(s), filename containing the R-code and a
                         date like 01-15-2014 (MM-DD-YYYY)

Matching is done by:
    1. R-code (the shared subfolder name) - primary key
    2. Study date - DICOM StudyDate (YYYYMMDD) vs the date parsed out of the
       PDF filename, to disambiguate when a folder holds more than one study.

Produces two output files:
    <output>.xlsx        - one row per study (R-code + study date): counts,
                            modalities, which PDF(s) it matched, and any
                            unmatched PDFs. Small - safe to eyeball in Excel.
    <output>_detail.csv  - one row per individual DICOM file, with the exact
                            PDF filename it was matched to (blank if none).
                            This is the file-level audit trail. It's a CSV,
                            not a second xlsx sheet, because Excel sheets cap
                            at ~1,048,576 rows - a million-image run would
                            blow past that - and CSV writes are streamed row
                            by row instead of held in memory.

Usage:
    python match_reports.py <images_folder> <reports_folder> [output.xlsx] [--debug]

    --debug prints, for every file/folder involved, exactly why it was or
    wasn't picked up: dcmread failures (with the exception message and
    whether forcing the read would have worked), parsed PDF dates, and a
    side-by-side of R-code folder names (with repr() to expose hidden
    whitespace/case differences) so a "should match" case can be diagnosed.
"""

import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

import pydicom
from openpyxl import Workbook

DATE_RE = re.compile(r"(\d{2})[-_.](\d{2})[-_.](\d{4})")


def parse_pdf_date(name):
    match = DATE_RE.search(name)
    if not match:
        return None
    month, day, year = match.groups()
    return f"{year}{month}{day}"


def read_dicom_study(path, debug=False):
    try:
        ds = pydicom.dcmread(str(path), stop_before_pixels=True, force=False)
    except Exception as exc:
        if debug:
            forced_ok = False
            forced_date = None
            try:
                ds2 = pydicom.dcmread(str(path), stop_before_pixels=True, force=True)
                forced_ok = True
                forced_date = str(ds2.get("StudyDate", ""))
            except Exception:
                pass
            hint = ""
            if forced_ok:
                hint = (
                    f" -> READABLE with force=True (StudyDate={forced_date!r}); "
                    "this file is likely missing the DICOM preamble/DICM header"
                )
            print(f"  [SKIP] {path}: {type(exc).__name__}: {exc}{hint}")
        return None
    study_date, modality = str(ds.get("StudyDate", "")), str(ds.get("Modality", ""))
    if debug:
        print(f"  [OK]   {path}: StudyDate={study_date!r} Modality={modality!r}")
    return study_date, modality


def scan_reports(reports_root, debug=False):
    """rcode -> {date_or_None: [pdf_path, ...]}"""
    reports = defaultdict(lambda: defaultdict(list))
    rcode_dirs = sorted(p for p in reports_root.iterdir() if p.is_dir())
    if debug:
        print(f"reports folder: found {len(rcode_dirs)} subfolders")
        for d in rcode_dirs:
            print(f"  rcode dir: {d.name!r}")
    for rcode_dir in rcode_dirs:
        rcode = rcode_dir.name
        if debug:
            print(f"\nScanning reports/{rcode}...")
        for path in sorted(rcode_dir.rglob("*.pdf")):
            if not path.is_file():
                continue
            date = parse_pdf_date(path.name)
            reports[rcode][date].append(path)
            if debug:
                print(f"  {path.name!r} -> parsed date {date!r}")
    return reports


def main():
    debug = "--debug" in sys.argv
    args = [a for a in sys.argv[1:] if a != "--debug"]

    if len(args) < 2:
        print(__doc__)
        sys.exit(1)

    images_root = Path(args[0]).expanduser().resolve()
    reports_root = Path(args[1]).expanduser().resolve()
    if not images_root.is_dir():
        print(f"Not a folder: {images_root}")
        sys.exit(1)
    if not reports_root.is_dir():
        print(f"Not a folder: {reports_root}")
        sys.exit(1)

    output_xlsx = Path(args[2]) if len(args) > 2 else Path.cwd() / "match_report.xlsx"
    detail_csv = output_xlsx.with_name(output_xlsx.stem + "_detail.csv")

    print("Scanning PDF reports...")
    reports = scan_reports(reports_root, debug=debug)

    if debug:
        print("\n--- R-code comparison ---")
        image_rcodes = {p.name for p in images_root.iterdir() if p.is_dir()}
        report_rcodes = set(reports)
        print(f"In both: {sorted(image_rcodes & report_rcodes)}")
        print(f"Only in images: {sorted(image_rcodes - report_rcodes)}")
        print(f"Only in reports: {sorted(report_rcodes - image_rcodes)}")
        print("--- end R-code comparison ---\n")

    print("Scanning DICOM images and writing per-file detail...")
    # (rcode, study_date) -> {"count": n, "modalities": set(), "matched_pdfs": set()}
    summary = defaultdict(lambda: {"count": 0, "modalities": set(), "matched_pdfs": set()})
    matched_pdf_paths = set()
    rcode_dirs = sorted(p for p in images_root.iterdir() if p.is_dir())

    with open(detail_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "R-Code", "DICOM File", "DICOM Study Date", "Modality",
            "Matched PDF File", "Match Status",
        ])

        for rcode_dir in rcode_dirs:
            rcode = rcode_dir.name
            rcode_reports = reports.get(rcode, {})
            if debug:
                print(f"\nScanning images/{rcode}...")
            for path in rcode_dir.rglob("*"):
                if not path.is_file():
                    continue
                info = read_dicom_study(path, debug=debug)
                if info is None:
                    continue
                study_date, modality = info

                bucket = summary[(rcode, study_date)]
                bucket["count"] += 1
                if modality:
                    bucket["modalities"].add(modality)

                candidate_pdfs = rcode_reports.get(study_date, []) if rcode_reports else []
                if candidate_pdfs:
                    status = "Matched (folder + date)"
                    pdf_names = "; ".join(p.name for p in candidate_pdfs)
                    for p in candidate_pdfs:
                        matched_pdf_paths.add(p)
                        bucket["matched_pdfs"].add(p.name)
                elif rcode_reports:
                    status = "Same folder, date unmatched"
                    pdf_names = ""
                else:
                    status = "No PDF report found"
                    pdf_names = ""

                writer.writerow([rcode, str(path), study_date, modality, pdf_names, status])

    print(f"Detail (per-image) report saved to: {detail_csv}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append([
        "R-Code", "Study Date", "DICOM File Count", "Modalities",
        "Matched PDF File(s)", "Match Status", "Other PDF Dates In Folder",
    ])

    matched = ambiguous = unmatched_images = 0
    for (rcode, study_date), info in sorted(summary.items()):
        rcode_reports = reports.get(rcode, {})
        other_dates = ", ".join(sorted(d or "unparsed" for d in rcode_reports if d != study_date))
        if info["matched_pdfs"]:
            status = "Matched (folder + date)"
            matched += 1
        elif rcode_reports:
            status = "Same folder, date unmatched - review manually"
            ambiguous += 1
        else:
            status = "No PDF report found"
            unmatched_images += 1
        ws.append([
            rcode, study_date, info["count"], ", ".join(sorted(info["modalities"])),
            ", ".join(sorted(info["matched_pdfs"])), status, other_dates,
        ])

    unmatched_pdf = 0
    for rcode, by_date in sorted(reports.items()):
        for date, pdf_list in sorted(by_date.items(), key=lambda kv: kv[0] or ""):
            for pdf_path in pdf_list:
                if pdf_path not in matched_pdf_paths:
                    ws.append([rcode, date or "", "", "", pdf_path.name, "No DICOM images found", ""])
                    unmatched_pdf += 1

    for column, width in zip("ABCDEFG", (12, 14, 16, 14, 40, 32, 26)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"
    wb.save(output_xlsx)

    print(f"\nMatched study-date buckets: {matched}")
    print(f"Same folder but date mismatch (needs review): {ambiguous}")
    print(f"DICOM studies with no PDF report: {unmatched_images}")
    print(f"PDF reports with no DICOM images: {unmatched_pdf}")
    print(f"Summary saved to: {output_xlsx}")
    print(f"Detail saved to: {detail_csv}")


if __name__ == "__main__":
    main()
