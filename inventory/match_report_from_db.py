"""Pass 3b - rebuild the old match_reports.py workbooks from the database.

match_reports.py answers one question - "for each study code, which DICOM
studies have a report and which don't?" - and writes it as one sheet per code
plus a combined summary. That output is what people downstream already read,
so this script reproduces it byte-for-byte in layout while getting its facts
from inventory.db instead of another walk of the share.

Nothing here opens the archive. crawl.py found the files, probe.py read the
DICOM study date and modality, report.py attributed everything to a code; all
that is left is the pairing, which is why this runs in seconds and can be
re-run as often as the matching rules need fixing.

The header, column widths, statuses and file naming are imported from
match_reports.py rather than copied, so the two cannot drift apart.

What differs from running match_reports.py directly
---------------------------------------------------
  * DICOM File Count comes from the studies table. With `probe.py
    --metadata all` every file was read and the count is exact; with
    `sample`/`none` some study counts are inferred from a sample, and this
    script prints how many rows that affects.
  * A study whose folder carries no code has nowhere to go in a per-code
    workbook, so it is left out unless you pass --include-unassigned. The
    master workbook's Unassigned sheet is the proper home for those.
  * Reports that no code could be derived for are likewise skipped; the count
    is printed.

Usage:
    python match_report_from_db.py --db inventory.db --out match_report.xlsx

    --db PATH        Database written by crawl.py / probe.py / report.py.
    --out PATH       Combined summary workbook (default match_report.xlsx).
    --results-dir D  Where the per-code workbooks go
                     (default: <out's folder>/results, as match_reports.py).
    --prefix LIST    Only codes with these prefixes, comma-separated
                     (e.g. --prefix AA). Default: all.
    --code LIST      Only these exact codes, comma-separated
                     (e.g. --code AA0006,AA0012). Default: all.
    --date-order X   Reading of ambiguous numeric dates: dmy (default) or mdy.
    --no-per-code    Write only the combined summary.
    --include-unassigned
                     Also emit studies whose folder carries no code, with an
                     empty Code column and the folder in place of the report.
"""

import argparse
import os
import re
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from match_reports import (  # noqa: E402  - reuse, don't reimplement
    SUMMARY_HEADER,
    display_modality,
    result_path,
    write_summary_sheet,
)
from report import parse_pdf_date, parse_pdf_modality  # noqa: E402

# The four strings match_reports.py puts in the Match Status column. Kept
# together because they are the contract with whatever reads these files.
STATUS_MATCHED = "Matched"
STATUS_NO_DATE_MATCH = "Same folder, no date/modality match - review manually"
STATUS_NO_REPORT = "No PDF report found"
STATUS_NO_IMAGES = "No DICOM images found"
STATUS_UNASSIGNED = "No code in folder path - unassigned"


def code_prefix(code):
    """The letters in front of the number: AA0006 -> AA."""
    m = re.match(r"[A-Za-z]+", code)
    return m.group(0).upper() if m else ""


def load_dirs(conn):
    """dir_id -> (path, parent_id, code). Small enough to hold in memory:
    one row per directory, not per file."""
    return {
        did: (path, parent, code)
        for did, path, parent, code in conn.execute(
            "SELECT id, path, parent_id, code FROM dirs"
        )
    }


def name_chain(dir_id, file_name, dirs):
    """The report's own name, then every folder name up to and including the
    code folder - the same chain match_reports.py searches when the file name
    alone carries no date or modality."""
    names = [file_name]
    code = dirs[dir_id][2]
    did = dir_id
    while did is not None and did in dirs:
        path, parent, _ = dirs[did]
        names.append(os.path.basename(path) or path)
        if code is None or parent not in dirs or dirs[parent][2] != code:
            break
        did = parent
    return names


def collect_reports(conn, dirs, date_order):
    """code -> [(full_path, date, modality)] for every PDF with a code."""
    reports = defaultdict(list)
    orphans = 0
    cur = conn.execute(
        "SELECT dir_id, name, code FROM files WHERE kind = 'pdf'"
    )
    for dir_id, name, code in cur:
        if not code:
            orphans += 1
            continue
        date = modality = None
        for part in name_chain(dir_id, name, dirs):
            if date is None:
                date = parse_pdf_date(part, date_order)
            if modality is None:
                modality = parse_pdf_modality(part)
            if date is not None and modality is not None:
                break
        full = os.path.join(dirs[dir_id][0], name)
        reports[code].append((full, date, modality))
    return reports, orphans


def build_indices(entries):
    """One code's reports -> (by (date, modality), by date) lookups."""
    by_date_modality = defaultdict(list)
    by_date = defaultdict(list)
    for full, date, modality in entries:
        if date is None:
            continue
        by_date[date].append(full)
        if modality is not None:
            by_date_modality[(date, modality)].append(full)
    return by_date_modality, by_date


def collect_studies(conn, dirs):
    """code -> {(date, modality): {"count", "inferred"}}.

    Studies are per-directory in the database but per (date, modality) in the
    old output, so several folders holding one study collapse into one row -
    which is what match_reports.py produced when it walked them together.
    """
    studies = defaultdict(lambda: defaultdict(
        lambda: {"count": 0, "inferred": False}))
    unassigned = []
    cur = conn.execute(
        "SELECT dir_id, study_date, modality, slice_count, confidence"
        " FROM studies"
    )
    for dir_id, date, modality, slices, confidence in cur:
        code = dirs[dir_id][2] if dir_id in dirs else None
        key = (date or "", modality or "")
        if not code:
            unassigned.append((dirs.get(dir_id, ("", None, None))[0], key,
                               slices or 0, confidence))
            continue
        bucket = studies[code][key]
        bucket["count"] += slices or 0
        if confidence != "counted":
            bucket["inferred"] = True
    return studies, unassigned


def rows_for_code(code, code_studies, code_reports):
    """One code's rows, in match_reports.py's order: studies first, sorted by
    date then modality, then any report nothing could be paired with."""
    by_date_modality, by_date = build_indices(code_reports)
    rows = []
    matched_paths = set()
    stats = defaultdict(int)

    for (date, modality), info in sorted(code_studies.items()):
        if modality and (date, modality) in by_date_modality:
            matches = by_date_modality[(date, modality)]
        elif date in by_date:
            matches = by_date[date]
        else:
            matches = []

        if matches:
            status = STATUS_MATCHED
            stats["matched"] += 1
            matched_paths.update(matches)
        elif code_reports:
            status = STATUS_NO_DATE_MATCH
            stats["ambiguous"] += 1
        else:
            status = STATUS_NO_REPORT
            stats["unmatched_images"] += 1
        if info["inferred"]:
            stats["inferred_rows"] += 1

        rows.append([
            code,
            date,
            display_modality(modality),
            info["count"],
            ", ".join(sorted(matches)),
            status,
        ])

    for full, date, modality in sorted(
        code_reports, key=lambda e: (e[1] or "", e[2] or "")
    ):
        if full in matched_paths:
            continue
        rows.append([
            code,
            date or "",
            display_modality(modality) or "",
            "",
            full,
            STATUS_NO_IMAGES,
        ])
        stats["unmatched_pdf"] += 1

    return rows, stats


def unassigned_rows(unassigned):
    """Studies with no code, folded into the same six columns: the folder
    stands in for the report, since that is the thing to go and look at."""
    grouped = defaultdict(lambda: {"count": 0, "folders": set()})
    for folder, key, slices, _confidence in unassigned:
        bucket = grouped[key]
        bucket["count"] += slices
        bucket["folders"].add(folder)
    return [
        ["", date, display_modality(modality), info["count"],
         ", ".join(sorted(info["folders"])), STATUS_UNASSIGNED]
        for (date, modality), info in sorted(grouped.items())
    ]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--out", default="match_report.xlsx")
    ap.add_argument("--results-dir", default=None)
    ap.add_argument("--prefix", default=None)
    ap.add_argument("--code", default=None)
    ap.add_argument("--date-order", default="dmy", choices=["dmy", "mdy"])
    ap.add_argument("--no-per-code", action="store_true")
    ap.add_argument("--include-unassigned", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    conn = sqlite3.connect(args.db)

    have = {r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type = 'table'")}
    if "studies" not in have:
        sys.exit("no studies table - run probe.py first")
    if not conn.execute(
            "SELECT 1 FROM files WHERE code IS NOT NULL LIMIT 1").fetchone():
        sys.exit("no codes assigned - run report.py --prefixes ... first")

    dirs = load_dirs(conn)
    reports, orphan_reports = collect_reports(conn, dirs, args.date_order)
    studies, unassigned = collect_studies(conn, dirs)
    conn.close()

    codes = set(studies) | set(reports)
    filtered = False
    if args.prefix:
        wanted = {p.strip().upper() for p in args.prefix.split(",") if p.strip()}
        codes = {c for c in codes if code_prefix(c) in wanted}
        filtered = True
        unknown = wanted - {code_prefix(c) for c in set(studies) | set(reports)}
        if unknown:
            print(f"warning: no codes with prefix {', '.join(sorted(unknown))} "
                  "- run report.py --discover to see what is there",
                  file=sys.stderr)
    if args.code:
        wanted = {c.strip().upper() for c in args.code.split(",") if c.strip()}
        codes = {c for c in codes if c.upper() in wanted}
        filtered = True

    out_path = Path(args.out).resolve()
    results_dir = Path(args.results_dir) if args.results_dir \
        else out_path.parent / "results"
    if not args.no_per_code:
        results_dir.mkdir(parents=True, exist_ok=True)

    combined = []
    totals = defaultdict(int)
    n_files = 0
    # Sorted by the per-code file name, so the combined sheet comes out in the
    # order match_reports.py's combine_results() produced by globbing them.
    for code in sorted(codes, key=lambda c: result_path(results_dir, c).name):
        rows, stats = rows_for_code(
            code, studies.get(code, {}), reports.get(code, []))
        if not rows:
            continue
        for key, value in stats.items():
            totals[key] += value
        combined.extend(rows)
        if not args.no_per_code:
            write_summary_sheet(result_path(results_dir, code), rows)
            n_files += 1

    if args.include_unassigned and not filtered:
        combined.extend(unassigned_rows(unassigned))

    write_summary_sheet(out_path, combined)

    print(f"wrote {out_path}  ({len(combined):,} rows)")
    if not args.no_per_code:
        print(f"  per-code workbooks: {n_files:,} in {results_dir}")
    print(f"  codes: {len(codes):,}   "
          f"matched: {totals['matched']:,}   "
          f"no date/modality match: {totals['ambiguous']:,}")
    print(f"  studies with no report: {totals['unmatched_images']:,}   "
          f"reports with no images: {totals['unmatched_pdf']:,}")
    if totals["inferred_rows"]:
        print(f"  NOTE: {totals['inferred_rows']:,} rows have a DICOM File "
              "Count inferred from a sample, not counted.\n"
              "        Re-run probe.py --metadata all for exact counts.")
    if unassigned and not args.include_unassigned:
        print(f"  skipped {len(unassigned):,} studies whose folder carries no "
              "code (--include-unassigned)")
    if orphan_reports:
        print(f"  skipped {orphan_reports:,} PDFs with no derivable code")


if __name__ == "__main__":
    main()
