"""Which patients have images, reports, both, or neither - and where a patient
is only partly covered, exactly which part is missing.

A patient is rarely all-or-nothing. AA0006 can have one study with a report,
a second study with none, and a third report whose images never arrived. A
single label per patient hides that, so every row carries the three counts
behind it and the label only says whether they are all accounted for:

    only image                   no report anywhere for this patient
    only report                  reported, but no DICOM file is attributed
    report and image (partial)   has both, but some study has no report, or
                                 some report has no images - the counts
                                 say which
    report and image (complete)  every study has a report and every report
                                 has images
    no report or image           a code folder holding neither - only
                                 Word/Excel files, stray litter, or nothing

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

IMAGES_ONLY = "only image"
REPORTS_ONLY = "only report"
PARTIAL = "report and image (partial)"
COMPLETE = "report and image (complete)"
NEITHER = "no report or image"

# Worst news first - the top of the list is the work queue.
BUCKETS = [IMAGES_ONLY, REPORTS_ONLY, PARTIAL, COMPLETE, NEITHER]

# What the caller counts per patient.
MATCHED = "report and image"
UNMATCHED = "only image"
ORPHAN = "only report"
UNDATED = "other pdf"
DETAIL = [MATCHED, UNMATCHED, ORPHAN, UNDATED]
# What a single row can be, in the order the breakdown lists them.
ROW_KINDS = DETAIL

# Every count column is named for the Status value it counts, so a number here
# and the rows behind it on the Studies sheet are found with the same phrase.
HEADER = ["Patient Code", "DICOM Files", "Report Files", "Other PDFs",
          "Studies", "Report and Image", "Only Image", "Only Report",
          "Category"]
WIDTHS = (14, 12, 12, 11, 9, 17, 12, 12, 26)


BREAKDOWN_HEADER = ["Patient Code", "Category", "Studies", "DICOM Files",
                    "Modalities", "Dates"]
BREAKDOWN_WIDTHS = (14, 20, 9, 12, 26, 60)

# A patient with 40 studies would otherwise push a cell past what anyone can
# read, and past Excel's 32k character limit at the extreme.
MAX_DATES = 12


def join_dates(dates):
    dates = sorted(d for d in dates if d)
    if not dates:
        return ""
    if len(dates) <= MAX_DATES:
        return ", ".join(dates)
    return ", ".join(dates[:MAX_DATES]) + f"  (+{len(dates) - MAX_DATES} more)"


def breakdown_rows(records):
    """One row per (patient, category), from (code, category, slices,
    modality, date) records.

    Coverage says a patient is partly covered; this says which studies are on
    which side of that - how many slices, what modalities, and on what dates -
    so the patient does not have to be looked up study by study.
    """
    groups = {}
    order = {name: i for i, name in enumerate(ROW_KINDS)}
    for code, category, slices, modality, date in records:
        g = groups.setdefault((code, category),
                              {"n": 0, "slices": 0, "mods": set(), "dates": set()})
        g["n"] += 1
        g["slices"] += slices or 0
        if modality:
            g["mods"].add(modality)
        if date:
            g["dates"].add(date)
    rows = [[code, category, g["n"], g["slices"],
             ", ".join(sorted(g["mods"])), join_dates(g["dates"])]
            for (code, category), g in groups.items()]
    rows.sort(key=lambda r: (r[0], order.get(r[1], 99)))
    return rows


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
        orphans, undated = d.get(ORPHAN, 0), d.get(UNDATED, 0)
        # A consent form or a manual is a PDF, not a report, so it must not
        # make a patient look reported.
        n_reports = max(n_pdf - undated, 0)
        bucket = bucket_of(n_dicom, n_reports, unmatched, orphans)
        summary[bucket] += 1
        totals[MATCHED] += matched
        totals[UNMATCHED] += unmatched
        totals[ORPHAN] += orphans
        totals[UNDATED] += undated
        rows.append([code, n_dicom, n_reports, undated, studies.get(code, 0),
                     matched, unmatched, orphans, bucket])
    rows.sort(key=lambda r: (order[r[8]], r[0]))
    return rows, summary, totals


def summary_lines(summary, totals=None):
    """The buckets, then the study-level totals, as (label, value) pairs."""
    lines = [(f"Patients: {name}", summary.get(name, 0)) for name in BUCKETS]
    lines.append(("Patients total", sum(summary.values())))
    if totals:
        lines.append(("", ""))
        lines += [(f"Rows: {name}", totals.get(name, 0)) for name in DETAIL]
    return lines


def print_summary(summary, totals=None, indent="  "):
    total = sum(summary.values())
    width = max(len(n) for n in BUCKETS + DETAIL) + 12
    for name in BUCKETS:
        n = summary.get(name, 0)
        pct = f"{100.0 * n / total:.1f}%" if total else "-"
        print(f"{indent}{'patients: ' + name:<{width}} {n:>8,}  {pct:>6}")
    print(f"{indent}{'patients total':<{width}} {total:>8,}")
    if totals:
        print()
        for name in DETAIL:
            print(f"{indent}{'rows: ' + name:<{width}} {totals.get(name, 0):>8,}")


def write_workbook(rows, summary, totals, path, breakdown=None):
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
    for letter, width in zip("ABCDEFGHI", WIDTHS):
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{ws.max_row}"

    if breakdown is not None:
        ws = wb.create_sheet("Breakdown")
        ws.append(BREAKDOWN_HEADER)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for row in breakdown:
            ws.append(row)
        for letter, width in zip("ABCDEF", BREAKDOWN_WIDTHS):
            ws.column_dimensions[letter].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:F{ws.max_row}"
    wb.save(path)
