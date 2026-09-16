"""Which patients have images, reports, both, or neither.

Four buckets, mutually exclusive, every attributed patient in exactly one:

    images, no report    images arrived, nothing has been reported yet
    report, no images    a report exists but the images are missing or
                         filed under a folder no code could be derived from
    images and report    complete
    neither              a code folder holding neither - only Word/Excel
                         files, stray litter, or nothing at all

Shared by report.py and match_report_from_db.py so the two always agree; the
counts come from the database, which means they cover every file crawled, not
only the studies probe.py could parse a header from.

"Has images" means at least one file identified as DICOM, not at least one
study. A folder whose headers could not be read still holds images, and
counting studies instead would quietly file it under "no images".
"""

from collections import Counter, defaultdict

IMAGES_ONLY = "images, no report"
REPORTS_ONLY = "report, no images"
BOTH = "images and report"
NEITHER = "neither"

# The order the buckets are reported in - worst news first, since the first
# two are the ones that need chasing.
BUCKETS = [IMAGES_ONLY, REPORTS_ONLY, BOTH, NEITHER]

HEADER = ["Patient Code", "DICOM Files", "Studies", "Report Files", "Category"]
WIDTHS = (14, 13, 9, 13, 20)


def bucket_of(n_dicom, n_pdf):
    if n_dicom and n_pdf:
        return BOTH
    if n_dicom:
        return IMAGES_ONLY
    if n_pdf:
        return REPORTS_ONLY
    return NEITHER


def classify(conn):
    """(rows, summary) - one row per patient, and the counts per bucket.

    rows are [code, dicom files, studies, pdf files, bucket], ordered by
    bucket then code so the work queue reads top-down.
    """
    counts = defaultdict(Counter)
    for code, kind, n in conn.execute(
            "SELECT code, kind, COUNT(*) FROM files"
            " WHERE code IS NOT NULL GROUP BY code, kind"):
        counts[code][kind or "unknown"] = n

    studies = dict(conn.execute(
        "SELECT d.code, COUNT(*) FROM studies s JOIN dirs d ON d.id = s.dir_id"
        " WHERE d.code IS NOT NULL GROUP BY d.code"))

    order = {name: i for i, name in enumerate(BUCKETS)}
    rows = []
    summary = Counter()
    for code, kinds in counts.items():
        n_dicom, n_pdf = kinds.get("dicom", 0), kinds.get("pdf", 0)
        bucket = bucket_of(n_dicom, n_pdf)
        summary[bucket] += 1
        rows.append([code, n_dicom, studies.get(code, 0), n_pdf, bucket])
    rows.sort(key=lambda r: (order[r[4]], r[0]))
    return rows, summary


def summary_lines(summary):
    """The four counts plus a total, as (label, value) pairs."""
    total = sum(summary.values())
    lines = [(f"Patients with {name}", summary.get(name, 0)) for name in BUCKETS]
    lines.append(("Patients total", total))
    return lines


def print_summary(summary, indent="  "):
    total = sum(summary.values())
    width = max(len(n) for n in BUCKETS) + 14
    for name in BUCKETS:
        n = summary.get(name, 0)
        pct = f"{100.0 * n / total:.1f}%" if total else "-"
        print(f"{indent}{'patients with ' + name:<{width}} {n:>8,}  {pct:>6}")
    print(f"{indent}{'patients total':<{width}} {total:>8,}")


def write_workbook(rows, summary, path):
    """The four buckets and the patients behind them, as its own workbook.

    Kept separate from match_report.xlsx on purpose: that file reproduces
    match_reports.py exactly, and anything reading it expects one sheet with
    six known columns.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Measure", "Patients"])
    ws["A1"].font = ws["B1"].font = Font(bold=True)
    for label, value in summary_lines(summary):
        ws.append([label, value])
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 12

    ws = wb.create_sheet("Patients")
    ws.append(HEADER)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    for letter, width in zip("ABCDE", WIDTHS):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    wb.save(path)
