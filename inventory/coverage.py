"""Which patients have images, reports, both, or neither - and where a patient
is only partly covered, exactly which part is missing.

A patient is rarely all-or-nothing. AA0006 can have one study with a report,
a second study with none, and a third report whose images never arrived. A
single label per patient hides that, so every row carries the three counts
behind it and the label only says whether they are all accounted for:

    images, no report    no report anywhere for this patient
    report, no images    reported, but no DICOM file is attributed to them
    partly covered       has both, but some study has no report, or some
                         report has no images - the counts say which
    fully covered        every study has a report and every report has images
    neither              a code folder holding neither - only Word/Excel
                         files, stray litter, or nothing at all

The per-study verdicts come from the caller, not from this module: report.py
and match_report_from_db.py match reports slightly differently, and each must
describe the workbook it is actually writing. File counts come from the
database, so they cover every file crawled rather than only the studies
probe.py could parse a header from.

"Has images" means at least one file identified as DICOM, not at least one
study. A folder whose headers could not be read still holds images, and
counting studies instead would quietly file it under "no images".
"""

from collections import Counter, defaultdict

IMAGES_ONLY = "images, no report"
REPORTS_ONLY = "report, no images"
PARTIAL = "partly covered"
COMPLETE = "fully covered"
NEITHER = "neither"

# Worst news first - the top of the list is the work queue.
BUCKETS = [IMAGES_ONLY, REPORTS_ONLY, PARTIAL, COMPLETE, NEITHER]

# What the caller counts per patient.
MATCHED = "studies with a report"
UNMATCHED = "studies with no report"
ORPHAN = "reports with no images"
DETAIL = [MATCHED, UNMATCHED, ORPHAN]

HEADER = ["Patient Code", "DICOM Files", "Report Files", "Studies",
          "Studies w/ Report", "Studies w/o Report", "Reports w/o Images",
          "Category"]
WIDTHS = (14, 12, 12, 9, 17, 18, 18, 18)


def new_detail():
    """code -> Counter of the three per-study verdicts."""
    return defaultdict(Counter)


def bucket_of(n_dicom, n_pdf, unmatched, orphans):
    if n_dicom and n_pdf:
        return PARTIAL if (unmatched or orphans) else COMPLETE
    if n_dicom:
        return IMAGES_ONLY
    if n_pdf:
        return REPORTS_ONLY
    return NEITHER


def classify(conn, detail=None):
    """(rows, summary, totals) for every attributed patient.

    detail is code -> Counter keyed by MATCHED / UNMATCHED / ORPHAN. Without
    it a patient holding both can only be called "partly covered", since
    nothing says whether the two sides line up.
    """
    detail = detail if detail is not None else {}
    counts = defaultdict(Counter)
    for code, kind, n in conn.execute(
            "SELECT code, kind, COUNT(*) FROM files"
            " WHERE code IS NOT NULL GROUP BY code, kind"):
        counts[code][kind or "unknown"] = n

    studies = dict(conn.execute(
        "SELECT d.code, COUNT(*) FROM studies s JOIN dirs d ON d.id = s.dir_id"
        " WHERE d.code IS NOT NULL GROUP BY d.code"))

    order = {name: i for i, name in enumerate(BUCKETS)}
    rows, summary, totals = [], Counter(), Counter()
    for code in set(counts) | set(detail):
        kinds = counts.get(code, Counter())
        n_dicom, n_pdf = kinds.get("dicom", 0), kinds.get("pdf", 0)
        d = detail.get(code, Counter())
        matched, unmatched = d.get(MATCHED, 0), d.get(UNMATCHED, 0)
        orphans = d.get(ORPHAN, 0)
        bucket = bucket_of(n_dicom, n_pdf, unmatched, orphans)
        summary[bucket] += 1
        totals[MATCHED] += matched
        totals[UNMATCHED] += unmatched
        totals[ORPHAN] += orphans
        rows.append([code, n_dicom, n_pdf, studies.get(code, 0),
                     matched, unmatched, orphans, bucket])
    rows.sort(key=lambda r: (order[r[7]], r[0]))
    return rows, summary, totals


def summary_lines(summary, totals=None):
    """The buckets, then the study-level totals, as (label, value) pairs."""
    lines = [(f"Patients {name}", summary.get(name, 0)) for name in BUCKETS]
    lines.append(("Patients total", sum(summary.values())))
    if totals:
        lines.append(("", ""))
        lines += [(f"Total {name}", totals.get(name, 0)) for name in DETAIL]
    return lines


def print_summary(summary, totals=None, indent="  "):
    total = sum(summary.values())
    width = max(len(n) for n in BUCKETS + DETAIL) + 14
    for name in BUCKETS:
        n = summary.get(name, 0)
        pct = f"{100.0 * n / total:.1f}%" if total else "-"
        print(f"{indent}{'patients ' + name:<{width}} {n:>8,}  {pct:>6}")
    print(f"{indent}{'patients total':<{width}} {total:>8,}")
    if totals:
        print()
        for name in DETAIL:
            print(f"{indent}{name:<{width}} {totals.get(name, 0):>8,}")


def write_workbook(rows, summary, totals, path):
    """The buckets and the patients behind them, as its own workbook.

    Kept separate from match_report.xlsx on purpose: that file reproduces
    match_reports.py exactly, and anything reading it expects one sheet with
    six known columns.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Measure", "Count"])
    ws["A1"].font = ws["B1"].font = Font(bold=True)
    for label, value in summary_lines(summary, totals):
        ws.append([label, value])
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12

    ws = wb.create_sheet("Patients")
    ws.append(HEADER)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for letter, width in zip("ABCDEFGH", WIDTHS):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{ws.max_row}"
    wb.save(path)
