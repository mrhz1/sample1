"""
Match PDF reports to DICOM studies.

Expects two folder trees sharing the same subfolder naming (e.g. "R123"):
    images/R123/...   -> DICOM files (mixed in with other junk, e.g. viewer apps)
    reports/R123/...  -> PDF report(s), filename containing the R-code and a
                         date like 01-15-2014 (MM-DD-YYYY)

Matching is done by:
    1. R-code (the shared subfolder name) - primary key
    2. Study date - DICOM StudyDate (YYYYMMDD) vs the date parsed out of the
       PDF filename, to disambiguate when a folder holds more than one study.

Usage:
    python match_reports.py <images_folder> <reports_folder> [output.xlsx] [--debug]

    --debug prints, for every file/folder involved, exactly why it was or
    wasn't picked up: dcmread failures (with the exception message and
    whether forcing the read would have worked), parsed PDF dates, and a
    side-by-side of R-code folder names (with repr() to expose hidden
    whitespace/case differences) so a "should match" case can be diagnosed.
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


def scan_images(images_root, debug=False):
    """rcode -> studydate -> {"count": n, "modalities": set()}"""
    studies = defaultdict(lambda: defaultdict(lambda: {"count": 0, "modalities": set()}))
    rcode_dirs = sorted(p for p in images_root.iterdir() if p.is_dir())
    if debug:
        print(f"images folder: found {len(rcode_dirs)} subfolders")
        for d in rcode_dirs:
            print(f"  rcode dir: {d.name!r}")
    for rcode_dir in rcode_dirs:
        rcode = rcode_dir.name
        if debug:
            print(f"\nScanning images/{rcode}...")
        for path in rcode_dir.rglob("*"):
            if not path.is_file():
                continue
            info = read_dicom_study(path, debug=debug)
            if info is None:
                continue
            study_date, modality = info
            bucket = studies[rcode][study_date]
            bucket["count"] += 1
            if modality:
                bucket["modalities"].add(modality)
        if debug:
            if studies.get(rcode):
                for study_date, info in sorted(studies[rcode].items()):
                    print(f"  -> study date {study_date!r}: {info['count']} file(s), modalities={sorted(info['modalities'])}")
            else:
                print("  -> no readable DICOM files found in this folder")
    return studies


def scan_reports(reports_root, debug=False):
    """rcode -> list of (pdf_path, parsed_date_or_None)"""
    reports = defaultdict(list)
    rcode_dirs = sorted(p for p in reports_root.iterdir() if p.is_dir())
    if debug:
        print(f"\nreports folder: found {len(rcode_dirs)} subfolders")
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
            reports[rcode].append((path, date))
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

    output = Path(args[2]) if len(args) > 2 else Path.cwd() / "match_report.xlsx"

    print("Scanning DICOM images...")
    studies = scan_images(images_root, debug=debug)
    print("Scanning PDF reports...")
    reports = scan_reports(reports_root, debug=debug)

    all_rcodes = sorted(set(studies) | set(reports))

    if debug:
        print("\n--- R-code comparison ---")
        only_images = set(studies) - set(reports)
        only_reports = set(reports) - set(studies)
        both = set(studies) & set(reports)
        print(f"In both: {sorted(both)}")
        print(f"Only in images: {sorted(only_images)}")
        print(f"Only in reports: {sorted(only_reports)}")
        print("--- end R-code comparison ---\n")

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
