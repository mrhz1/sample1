"""
Match PDF reports to DICOM studies.

Expects two folder trees sharing the same subfolder naming (e.g. "R123"):
    images/R123/...   -> DICOM files (mixed in with other junk, e.g. viewer apps)
    reports/R123/...  -> PDF report(s), filename containing the R-code and a
                         date like 07-07-2007 (DD-MM-YYYY)

Matching is done by:
    1. R-code (the shared subfolder name) - primary key
    2. Study date - DICOM StudyDate (YYYYMMDD) vs the date parsed out of the
       PDF filename, to disambiguate when a folder holds more than one study.

Usage:
    python match_reports.py <images_folder> <reports_folder> [output.xlsx]
"""

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
    day, month, year = match.groups()
    return f"{year}{month}{day}"


def read_dicom_study(path):
    try:
        ds = pydicom.dcmread(str(path), stop_before_pixels=True, force=False)
    except Exception:
        return None
    return str(ds.get("StudyDate", "")), str(ds.get("Modality", ""))


def scan_images(images_root):
    """rcode -> studydate -> {"count": n, "modalities": set()}"""
    studies = defaultdict(lambda: defaultdict(lambda: {"count": 0, "modalities": set()}))
    for rcode_dir in sorted(p for p in images_root.iterdir() if p.is_dir()):
        rcode = rcode_dir.name
        for path in rcode_dir.rglob("*"):
            if not path.is_file():
                continue
            info = read_dicom_study(path)
            if info is None:
                continue
            study_date, modality = info
            bucket = studies[rcode][study_date]
            bucket["count"] += 1
            if modality:
                bucket["modalities"].add(modality)
    return studies


def scan_reports(reports_root):
    """rcode -> list of (pdf_path, parsed_date_or_None)"""
    reports = defaultdict(list)
    for rcode_dir in sorted(p for p in reports_root.iterdir() if p.is_dir()):
        rcode = rcode_dir.name
        for path in sorted(rcode_dir.rglob("*.pdf")):
            if not path.is_file():
                continue
            reports[rcode].append((path, parse_pdf_date(path.name)))
    return reports


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    images_root = Path(sys.argv[1]).expanduser().resolve()
    reports_root = Path(sys.argv[2]).expanduser().resolve()
    if not images_root.is_dir():
        print(f"Not a folder: {images_root}")
        sys.exit(1)
    if not reports_root.is_dir():
        print(f"Not a folder: {reports_root}")
        sys.exit(1)

    output = Path(sys.argv[3]) if len(sys.argv) > 3 else Path.cwd() / "match_report.xlsx"

    print("Scanning DICOM images...")
    studies = scan_images(images_root)
    print("Scanning PDF reports...")
    reports = scan_reports(reports_root)

    all_rcodes = sorted(set(studies) | set(reports))

    wb = Workbook()
    ws = wb.active
    ws.title = "Matches"
    ws.append([
        "R-Code", "PDF File", "PDF Date", "DICOM Study Date",
        "DICOM File Count", "Modalities", "Match Status",
    ])

    matched = unmatched_pdf = unmatched_images = ambiguous = 0

    for rcode in all_rcodes:
        rcode_studies = studies.get(rcode, {})
        rcode_reports = reports.get(rcode, [])

        if not rcode_reports and rcode_studies:
            for study_date, info in sorted(rcode_studies.items()):
                ws.append([
                    rcode, "", "", study_date, info["count"],
                    ", ".join(sorted(info["modalities"])), "No PDF report found",
                ])
                unmatched_images += 1
            continue

        if rcode_reports and not rcode_studies:
            for pdf_path, pdf_date in rcode_reports:
                ws.append([
                    rcode, pdf_path.name, pdf_date or "", "", "", "",
                    "No DICOM images found",
                ])
                unmatched_pdf += 1
            continue

        for pdf_path, pdf_date in rcode_reports:
            if pdf_date and pdf_date in rcode_studies:
                info = rcode_studies[pdf_date]
                ws.append([
                    rcode, pdf_path.name, pdf_date, pdf_date, info["count"],
                    ", ".join(sorted(info["modalities"])), "Matched (folder + date)",
                ])
                matched += 1
            else:
                candidate_dates = ", ".join(sorted(rcode_studies))
                ws.append([
                    rcode, pdf_path.name, pdf_date or "", candidate_dates,
                    sum(i["count"] for i in rcode_studies.values()),
                    ", ".join(sorted({m for i in rcode_studies.values() for m in i["modalities"]})),
                    "Same folder, date unmatched - review manually",
                ])
                ambiguous += 1

    for column, width in zip("ABCDEFG", (12, 45, 12, 30, 16, 14, 32)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"

    wb.save(output)
    print(f"\nMatched: {matched}")
    print(f"Same folder but date mismatch (needs review): {ambiguous}")
    print(f"PDF with no DICOM images: {unmatched_pdf}")
    print(f"DICOM images with no PDF report: {unmatched_images}")
    print(f"Report saved to: {output}")


if __name__ == "__main__":
    main()
