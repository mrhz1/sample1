"""
Match PDF reports to DICOM studies and record the exact file-level linkage.

Expects two folder trees sharing the same subfolder naming (e.g. "R123"):
    images/R123/...   -> DICOM files (mixed in with other junk, e.g. viewer apps)
    reports/R123/...  -> PDF report(s), filename containing the R-code, a
                         modality (CT / MRI / Echo / ...), and a date like
                         01-15-2014 (MM-DD-YYYY), e.g. "R123 CT 10-10-2010.pdf"

Matching is done in tiers, strongest first:
    1. R-code (the shared subfolder name) - required for any match at all.
    2. Study date + Modality - DICOM StudyDate (YYYYMMDD) and Modality
       (CT/MR/US/...) both matched against the PDF filename. This is what
       correctly splits same-day-different-study cases (e.g. a CT and an
       Echo done on the same date) that date-only matching would conflate.
    3. Study date only - used as a fallback when the modality can't be
       parsed from the PDF filename or is missing from the DICOM tag.
    Anything weaker than that is left for manual review rather than guessed.

Produces two output files:
    <output>.xlsx         - one row per study (R-code + date + modality):
                             counts, which PDF(s) it matched, and any
                             unmatched PDFs. Small - safe to eyeball in Excel.
    <output>_detail.xlsx  - one row per individual DICOM file, with the exact
                             PDF full path it was matched to (blank if none).
                             This is the file-level audit trail, written with
                             openpyxl's write-only/streaming mode to keep
                             memory flat regardless of file count. NOTE: an
                             Excel sheet caps at ~1,048,576 rows - a
                             many-million-image run WILL exceed that and
                             openpyxl will raise when it happens. If your
                             real dataset is that large, ask for the CSV
                             variant back (or a split-by-rcode xlsx) instead.

    Modality is shown as a human-readable label (e.g. "MRI" instead of the
    raw DICOM code "MR") in both files - matching logic still uses the raw
    DICOM codes internally, this only affects what's displayed.

Usage:
    python match_reports.py <images_folder> <reports_folder> [output.xlsx] [--debug]

    --debug prints, for every file/folder involved, exactly why it was or
    wasn't picked up: dcmread failures (with the exception message and
    whether forcing the read would have worked), parsed PDF dates/modality,
    and a side-by-side of R-code folder names (with repr() to expose hidden
    whitespace/case differences) so a "should match" case can be diagnosed.
"""

import re
import sys
from collections import defaultdict
from pathlib import Path

import pydicom
from openpyxl import Workbook

DATE_RE = re.compile(r"(\d{2})[-_.](\d{2})[-_.](\d{4})")

# PDF-filename token (uppercased) -> standard DICOM Modality code.
# Extend this as you find more variants in the real filenames.
MODALITY_ALIASES = {
    "CT": "CT",
    "MRI": "MR",
    "MR": "MR",
    "ECHO": "US",
    "ECHOCARDIOGRAM": "US",
    "ULTRASOUND": "US",
    "US": "US",
    "XRAY": "CR",
    "X-RAY": "CR",
    "DX": "DX",
    "MAMMO": "MG",
    "MG": "MG",
    "PET": "PT",
    "NM": "NM",
}
MODALITY_RE = re.compile(
    r"(?<![A-Z0-9])(" + "|".join(re.escape(k) for k in MODALITY_ALIASES) + r")(?![A-Z0-9])"
)

# DICOM Modality code -> human-readable label, for display in the Excel
# outputs only. Extend as you encounter more codes in the real data.
MODALITY_DISPLAY_NAMES = {
    "CT": "CT",
    "MR": "MRI",
    "US": "Echo",
    "CR": "X-Ray",
    "DX": "X-Ray",
    "MG": "Mammography",
    "PT": "PET",
    "NM": "Nuclear Medicine",
    "XA": "Angiography",
    "RF": "Fluoroscopy",
    "SR": "Structured Report",
}


def display_modality(code):
    if not code:
        return code
    return MODALITY_DISPLAY_NAMES.get(code, code)


def parse_pdf_date(name):
    match = DATE_RE.search(name)
    if not match:
        return None
    month, day, year = match.groups()
    return f"{year}{month}{day}"


def parse_pdf_modality(name):
    match = MODALITY_RE.search(name.upper())
    if not match:
        return None
    return MODALITY_ALIASES[match.group(1)]


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
    """rcode -> list of {"path", "date", "modality"}"""
    reports = defaultdict(list)
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
            modality = parse_pdf_modality(path.name)
            reports[rcode].append({"path": path, "date": date, "modality": modality})
            if debug:
                print(f"  {path.name!r} -> date={date!r} modality={modality!r}")
    return reports


def build_indices(rcode_reports):
    """entries -> (by_date_modality, by_date) lookup dicts for one R-code."""
    by_date_modality = defaultdict(list)
    by_date = defaultdict(list)
    for entry in rcode_reports:
        if entry["date"] is None:
            continue
        by_date[entry["date"]].append(entry)
        if entry["modality"] is not None:
            by_date_modality[(entry["date"], entry["modality"])].append(entry)
    return by_date_modality, by_date


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
    detail_xlsx = output_xlsx.with_name(output_xlsx.stem + "_detail.xlsx")

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
    # (rcode, study_date, modality) -> {"count", "matched_pdfs"}
    summary = defaultdict(lambda: {"count": 0, "matched_pdfs": set()})
    matched_pdf_paths = set()
    rcode_dirs = sorted(p for p in images_root.iterdir() if p.is_dir())

    detail_wb = Workbook(write_only=True)
    detail_ws = detail_wb.create_sheet(title="Detail")
    for column, width in zip("ABCDEF", (12, 60, 14, 12, 60, 40)):
        detail_ws.column_dimensions[column].width = width
    detail_ws.freeze_panes = "A2"
    detail_ws.append([
        "R-Code", "DICOM File", "DICOM Study Date", "Modality",
        "Matched PDF File", "Match Status",
    ])

    for rcode_dir in rcode_dirs:
        rcode = rcode_dir.name
        rcode_reports = reports.get(rcode, [])
        by_date_modality, by_date = build_indices(rcode_reports)
        if debug:
            print(f"\nScanning images/{rcode}...")
        for path in rcode_dir.rglob("*"):
            if not path.is_file():
                continue
            info = read_dicom_study(path, debug=debug)
            if info is None:
                continue
            study_date, modality = info

            if modality and (study_date, modality) in by_date_modality:
                matches = by_date_modality[(study_date, modality)]
                status = "Matched (folder + date + modality)"
            elif study_date in by_date:
                matches = by_date[study_date]
                status = "Matched (folder + date, modality unconfirmed)"
            else:
                matches = []
                status = "Same folder, date unmatched" if rcode_reports else "No PDF report found"

            bucket = summary[(rcode, study_date, modality)]
            bucket["count"] += 1

            pdf_paths_str = "; ".join(str(m["path"]) for m in matches)
            for m in matches:
                matched_pdf_paths.add(m["path"])
                bucket["matched_pdfs"].add(str(m["path"]))

            detail_ws.append([rcode, str(path), study_date, display_modality(modality), pdf_paths_str, status])

    detail_wb.save(detail_xlsx)
    print(f"Detail (per-image) report saved to: {detail_xlsx}")

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append([
        "R-Code", "Study Date", "Modality", "DICOM File Count",
        "Matched PDF File(s)", "Match Status", "Other PDFs In Folder",
    ])

    matched = ambiguous = unmatched_images = 0
    for (rcode, study_date, modality), info in sorted(summary.items()):
        rcode_reports = reports.get(rcode, [])
        others = ", ".join(
            f"{e['date'] or 'unparsed'}/{display_modality(e['modality']) or '?'}"
            for e in rcode_reports
            if str(e["path"]) not in info["matched_pdfs"]
        )
        if info["matched_pdfs"]:
            status = "Matched"
            matched += 1
        elif rcode_reports:
            status = "Same folder, no date/modality match - review manually"
            ambiguous += 1
        else:
            status = "No PDF report found"
            unmatched_images += 1
        ws.append([
            rcode, study_date, display_modality(modality), info["count"],
            ", ".join(sorted(info["matched_pdfs"])), status, others,
        ])

    unmatched_pdf = 0
    for rcode, entries in sorted(reports.items()):
        for entry in sorted(entries, key=lambda e: (e["date"] or "", e["modality"] or "")):
            if entry["path"] not in matched_pdf_paths:
                ws.append([
                    rcode, entry["date"] or "", display_modality(entry["modality"]) or "", "",
                    str(entry["path"]), "No DICOM images found", "",
                ])
                unmatched_pdf += 1

    for column, width in zip("ABCDEFG", (12, 14, 10, 16, 45, 40, 26)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"
    wb.save(output_xlsx)

    print(f"\nMatched study buckets: {matched}")
    print(f"Same folder, no date/modality match (needs review): {ambiguous}")
    print(f"DICOM studies with no PDF report: {unmatched_images}")
    print(f"PDF reports with no DICOM images: {unmatched_pdf}")
    print(f"Summary saved to: {output_xlsx}")
    print(f"Detail saved to: {detail_xlsx}")


if __name__ == "__main__":
    main()
