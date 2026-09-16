"""Pass 3 - attribute files to patients, match reports, write the master Excel.

Reads only the database, never the archive. Pass 2 already extracted and stored
the DICOM study date and modality, so matching here is a join rather than a
scan - which is why this runs in seconds and can be re-run freely. That matters
because the code-matching rules are the part that will be wrong on the first
try against a messy archive.

Date and modality parsing is imported from match_reports.py rather than
reimplemented, so the two tools stay in step as MODALITY_ALIASES grows.

How a file gets attributed to a patient
---------------------------------------
The DICOM headers carry no patient identifier, so the path is the only link.
Evidence is taken from the most specific source available:

  1. the file's own name, if it contains a code
  2. otherwise the nearest ancestor folder whose name contains a code

When the two disagree the file is still assigned (by --prefer, default
'filename') and also listed on the Conflicts sheet, because one of them is a
filing error and only a person can say which.

Codes are normalised on the number, not the text: AA1, AA001 and AA0001 are one
patient. Padding is learned per prefix, since AVDD001 and AA0001 don't agree on
width and a single global setting would corrupt one of them.

The results are written back to the database (file_code, dir_code) so you can
query them directly - see README.md for examples.

Usage:
    python report.py --db inventory.db --discover
    python report.py --db inventory.db --prefixes AA,AVDD,QQQ --out master.xlsx

    --discover       List every code-like pattern in the archive with counts,
                     so you can see which prefixes are real before committing
                     to them. Writes nothing.
    --prefixes LIST  Comma-separated study prefixes to treat as patient codes.
    --code-regex RE  Full override, if --prefixes isn't expressive enough.
    --date-order X   Reading of ambiguous numeric dates: dmy (default) or mdy.
    --pad SPEC       Force the display width of the number instead of learning
                     it. '4' applies to every prefix, 'AA=4,AVDD=3' per prefix.
                     Use when one junk name has widened a whole prefix - see
                     diagnose_codes.py.
    --prefer WHICH   On a filename/folder conflict, trust 'filename' (default)
                     or 'folder'.
    --csv PATH       Also dump the file-level inventory to CSV. Millions of
                     rows - it will not fit in Excel, hence a separate file.
"""

import argparse
import csv
import inspect
import os
import re
import sqlite3
import sys
import time
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from match_reports import (  # noqa: E402  - reuse, don't reimplement
    display_modality,
    parse_pdf_date as _parse_pdf_date,
    parse_pdf_modality,
)

# Older copies of match_reports.py have parse_pdf_date(name) with no
# date_order. Adapt rather than require a particular vintage of that file, but
# say so loudly - silently ignoring --date-order would misread every date in a
# mdy archive.
if len(inspect.signature(_parse_pdf_date).parameters) >= 2:
    parse_pdf_date = _parse_pdf_date
    HAS_DATE_ORDER = True
else:
    HAS_DATE_ORDER = False

    def parse_pdf_date(name, date_order="dmy"):
        return _parse_pdf_date(name)
import coverage  # noqa: E402
from probe import KNOWN_EXTS  # noqa: E402

KIND_COLUMNS = ["dicom", "dicomdir", "pdf", "word", "excel", "slides", "image",
                "video", "archive", "program", "text", "office", "unknown",
                "other"]

# Anything shaped like a code, used only by --discover. Deliberately loose: it
# will match "batch 1" too, which is the point - you look at the counts and
# decide what is real rather than trusting a guess.
CANDIDATE_RE = re.compile(r"\b([A-Za-z]{2,10})[-_ ]?(\d+)\b")

# Codes live on files/dirs themselves (see crawl.py), so pass 3 creates no
# tables of its own - only the index, once the column is populated.
INDEX = "CREATE INDEX IF NOT EXISTS files_code ON files(code);"

# Convenience view: adds the folder path, which lives on dirs.
VIEW = """
DROP VIEW IF EXISTS v_files;
CREATE VIEW v_files AS
SELECT f.id           AS file_id,
       f.code         AS code,
       f.kind         AS kind,
       f.code_source  AS code_source,
       f.name         AS name,
       f.ext          AS ext,
       f.size         AS size,
       d.path         AS folder,
       d.path || '/' || f.name AS full_path
  FROM files f
  JOIN dirs d ON d.id = f.dir_id;
"""


def discover(conn):
    """Report every code-like token in the archive, with counts and examples.

    There is no safe default regex when prefixes vary by study and new ones
    turn up later, so this exists to let you look before choosing.
    """
    hits = defaultdict(Counter)
    examples = {}
    for (path,) in conn.execute("SELECT path FROM dirs"):
        for part in path.split(os.sep):
            for prefix, digits in CANDIDATE_RE.findall(part):
                p = prefix.upper()
                hits[p][len(digits)] += 1
                examples.setdefault(p, part)
    for (name,) in conn.execute("SELECT name FROM files"):
        for prefix, digits in CANDIDATE_RE.findall(name):
            p = prefix.upper()
            hits[p][len(digits)] += 1
            examples.setdefault(p, name)

    print(f"{'PREFIX':<12} {'COUNT':>9}  {'DIGIT WIDTHS':<26} EXAMPLE")
    print("-" * 88)
    for prefix, widths in sorted(hits.items(), key=lambda kv: -sum(kv[1].values())):
        total = sum(widths.values())
        w = ", ".join(f"{k}({v:,})" for k, v in sorted(widths.items()))
        print(f"{prefix:<12} {total:>9,}  {w:<26} {examples[prefix][:32]}")
    print("\nPick the real study prefixes and pass them, e.g.:")
    print("  python report.py --db inventory.db --prefixes AA,AVDD,QQQ")
    print("\nMixed digit widths for one prefix are fine and expected - they are")
    print("normalised on the number, so AA001 and AA0001 are one patient.")


def build_code_re(prefixes, override):
    if override:
        return re.compile(override, re.IGNORECASE)
    if not prefixes:
        sys.exit("give --prefixes (run --discover first) or --code-regex")
    alt = "|".join(re.escape(p.strip()) for p in prefixes.split(",") if p.strip())
    # (?!\d) matters more than it looks. Without it "AA 20240115 rescan" matches
    # as AA20240 - a patient who does not exist, whose 5-digit number then pads
    # every real AA0001 to AA00001. A code is the whole digit run or nothing.
    return re.compile(rf"(?:{alt})[-_ ]?\d{{1,5}}(?!\d)", re.IGNORECASE)


def parse_code(text, code_re):
    """Return (PREFIX, number, digit_width) or None.

    Keyed on the integer, so padding differences can't split one patient into
    several. Width is tracked only to format the code back for display.
    """
    m = code_re.search(text)
    if not m:
        return None
    raw = m.group(0)
    dm = re.search(r"\d+", raw)
    prefix = raw[: dm.start()].strip(" -_").upper()
    return prefix, int(dm.group(0)), len(dm.group(0))


def attribute(conn, code_re, prefer, root):
    """Assign a code to every directory and file; write both back to the DB."""
    conn.execute("UPDATE dirs SET code = NULL WHERE code IS NOT NULL")
    conn.execute("UPDATE files SET code = NULL, prefix = NULL, number = NULL,"
                 " code_source = NULL, conflict_with = NULL"
                 " WHERE code_source IS NOT NULL")

    dir_paths = dict(conn.execute("SELECT id, path FROM dirs"))
    widths = defaultdict(Counter)

    dir_code = {}
    for dir_id, path in dir_paths.items():
        rel = os.path.relpath(path, root) if path != root else ""
        found = None
        for part in reversed(rel.split(os.sep)):
            found = parse_code(part, code_re)
            if found:
                break
        if found:
            prefix, number, w = found
            widths[prefix][w] += 1
            dir_code[dir_id] = (prefix, number)
        else:
            dir_code[dir_id] = None

    # Cheapest possible pre-filter: a name can only hold a code if it contains
    # one of the prefixes. Skips the regex on millions of DICOM slice names.
    hints = tuple(re.findall(r"[A-Za-z]{2,10}", code_re.pattern))

    def rows():
        cur = conn.execute("SELECT id, dir_id, name, ext, size FROM files")
        for fid, did, name, ext, size in cur:
            folder = dir_code.get(did)
            fname = None
            upper = name.upper()
            if any(h.upper() in upper for h in hints):
                found = parse_code(name, code_re)
                if found:
                    prefix, number, w = found
                    widths[prefix][w] += 1
                    fname = (prefix, number)
            if fname and folder and fname != folder:
                code = fname if prefer == "filename" else folder
                conflict = folder if prefer == "filename" else fname
                source = f"{prefer} (conflict)"
            elif fname:
                code, conflict, source = fname, None, "filename"
            elif folder:
                code, conflict, source = folder, None, "folder"
            else:
                code, conflict, source = None, None, "none"
            yield fid, did, name, ext, size, code, source, conflict

    return dir_code, dir_paths, rows, widths


def resolve_widths(counts):
    """prefix -> the digit width to display, given every width seen for it.

    The most common one, not the widest. A handful of odd names cannot then
    repad thousands of real codes, and nothing is lost when they are genuine:
    fmt() pads to this width but never truncates, so a longer number than
    expected still prints in full.
    """
    return {prefix: max(seen.items(), key=lambda kv: (kv[1], kv[0]))[0]
            for prefix, seen in counts.items()}


def apply_pad(widths, spec):
    """Override the learned padding width.

    Widths are otherwise the widest digit run the pattern matched for that
    prefix, so a single name like "AA 20240115 rescan" pads all 1,400 real
    codes to 5 digits. This is the escape hatch when that happens.
    """
    if not spec:
        return
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            prefix, _, value = part.partition("=")
            prefix = prefix.strip().upper()
            if prefix not in widths:
                print(f"note: --pad names prefix {prefix}, which was not found",
                      file=sys.stderr)
            widths[prefix] = int(value)
        else:
            for prefix in list(widths):
                widths[prefix] = int(part)


def fmt(code, widths):
    if not code:
        return ""
    prefix, number = code
    return f"{prefix}{number:0{max(widths.get(prefix, 1), 1)}d}"


def kind_of(ext, probed):
    if probed:
        return probed
    return KNOWN_EXTS.get(ext, "other" if ext else "unknown")


def build(args):
    conn = sqlite3.connect(args.db)
    root = conn.execute("SELECT value FROM meta WHERE key='root'").fetchone()[0]
    code_re = build_code_re(args.prefixes, args.code_regex)

    probed_kind = dict(conn.execute(
        "SELECT id, kind FROM files WHERE kind IS NOT NULL"))
    dir_code, dir_paths, rows, widths = attribute(conn, code_re, args.prefer, root)

    per_patient = defaultdict(Counter)
    per_patient_bytes = Counter()
    conflicts, orphan_reports = [], []
    reports = defaultdict(list)
    unassigned_dirs = Counter()
    kind_totals = Counter()
    pending = []

    csv_fh = csv_writer = None
    if args.csv:
        csv_fh = open(args.csv, "w", newline="", encoding="utf-8")
        csv_writer = csv.writer(csv_fh)
        csv_writer.writerow(["code", "code_source", "kind", "size", "path"])

    n_total = conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
    n_seen = 0
    t_started = t_last = time.time()

    for fid, did, name, ext, size, code, source, conflict in rows():
        n_seen += 1
        if n_seen % 50000 == 0:
            now = time.time()
            if now - t_last >= 5.0:
                rate = n_seen / max(now - t_started, 0.001)
                eta = (n_total - n_seen) / rate / 60 if rate else 0
                print(f"  {n_seen:,}/{n_total:,} files "
                      f"({100.0 * n_seen / max(n_total,1):.1f}%)  "
                      f"eta~{eta:.1f} min", flush=True)
                t_last = now
        kind = kind_of(ext, probed_kind.get(fid))
        kind_totals[kind] += 1
        full = os.path.join(dir_paths[did], name)

        key = code or "UNASSIGNED"
        per_patient[key][kind] += 1
        per_patient_bytes[key] += size or 0
        if not code:
            unassigned_dirs[did] += 1
        if conflict:
            conflicts.append((code, conflict, source, full))

        if kind == "pdf":
            folder_name = os.path.basename(dir_paths[did])
            date = (parse_pdf_date(name, args.date_order)
                    or parse_pdf_date(folder_name, args.date_order))
            modality = parse_pdf_modality(name) or parse_pdf_modality(folder_name)
            entry = (full, date, modality, name)
            if code:
                reports[code].append(entry)
            else:
                orphan_reports.append(entry)

        pending.append((fid, code[0] if code else None,
                        code[1] if code else None, source,
                        conflict, kind))
        if len(pending) >= 20000:
            flush(conn, pending)
            pending = []
        if csv_writer:
            csv_writer.writerow([code, source, kind, size or 0, full])

    flush(conn, pending)
    if csv_fh:
        csv_fh.close()

    widths = resolve_widths(widths)
    apply_pad(widths, args.pad)

    # Padding is only known now that everything has been seen, so fill the
    # display column in one statement per prefix.
    for prefix, width in widths.items():
        conn.execute(
            f"UPDATE files SET code = prefix || "
            f"substr('00000000', 1, {width} - length(CAST(number AS TEXT))) "
            f"|| CAST(number AS TEXT) WHERE prefix = ?", (prefix,))
    conn.executemany(
        "UPDATE dirs SET code = ? WHERE id = ?",
        [(fmt(c, widths) or None, d) for d, c in dir_code.items()])
    conn.execute(INDEX)
    conn.executescript(VIEW)
    conn.commit()

    studies = match_studies(conn, dir_code, dir_paths, reports, args)
    write_matches(conn, studies, widths)
    suggestions = suggest(studies, reports)
    return dict(
        root=root, per_patient=per_patient, per_patient_bytes=per_patient_bytes,
        conflicts=conflicts, studies=studies, suggestions=suggestions,
        unassigned_dirs=unassigned_dirs, dir_paths=dir_paths,
        kind_totals=kind_totals, reports=reports, widths=widths,
        orphan_reports=orphan_reports, conn=conn)


def flush(conn, pending):
    if pending:
        conn.executemany(
            "UPDATE files SET prefix = ?, number = ?, code_source = ?,"
            " conflict_with = ?, kind = COALESCE(kind, ?),"
            " kind_source = COALESCE(kind_source, 'ext') WHERE id = ?",
            [(p, n, s, str(c) if c else None, k, f)
             for f, p, n, s, c, k in pending])


def match_studies(conn, dir_code, dir_paths, reports, args):
    """Pair each study with a report, strongest evidence first."""
    out, used = [], defaultdict(set)
    rows = conn.execute(
        "SELECT id, dir_id, study_uid, study_date, modality, slice_count,"
        " confidence FROM studies ORDER BY dir_id").fetchall()

    for sid, did, uid, date, modality, slices, confidence in rows:
        code = dir_code.get(did)
        cands = reports.get(code, []) if code else []
        matched, status = None, "no report found"
        if not code:
            status = "unassigned - no code in path"
        else:
            exact = [r for r in cands if r[1] and r[1] == date
                     and r[2] and modality and r[2] == modality]
            same_day = [r for r in cands if r[1] and r[1] == date]
            if exact:
                matched, status = exact[0], "matched (date + modality)"
            elif same_day:
                matched = same_day[0]
                status = "matched (date only - modality unconfirmed)"
            elif cands:
                status = "report exists for patient but no date match"
        if matched:
            used[code].add(matched[0])
        out.append(dict(
            study_id=sid, code=code, dir=dir_paths[did], date=date,
            modality=display_modality(modality) if modality else "",
            raw_modality=modality or "", slices=slices, confidence=confidence,
            report=os.path.basename(matched[0]) if matched else "",
            status=status))

    for code, items in reports.items():
        for path, date, modality, name in items:
            if path not in used[code]:
                out.append(dict(
                    study_id=None, code=code, dir=os.path.dirname(path),
                    date=date or "",
                    modality=display_modality(modality) if modality else "",
                    raw_modality=modality or "", slices=0, confidence="",
                    report=name, status="report with no matching images"))

    return out


MATCH_SCHEMA = """
DROP TABLE IF EXISTS study_report;
CREATE TABLE study_report (
    study_id    INTEGER,        -- NULL for a report with no images
    code        TEXT,
    study_date  TEXT,
    modality    TEXT,           -- raw DICOM code: MR, US, CR
    slices      INTEGER,        -- DICOM files in the study
    report      TEXT,           -- report file name, '' if none
    status      TEXT NOT NULL,
    folder      TEXT
);
CREATE INDEX study_report_code   ON study_report(code);
CREATE INDEX study_report_status ON study_report(status);
"""


def write_matches(conn, out, widths):
    """Mirror the match result into the database so it can be queried.

    Codes are held internally as (prefix, number); store the display form so
    the table can be joined against files.code without conversion.
    """
    conn.executescript(MATCH_SCHEMA)
    conn.executemany(
        "INSERT INTO study_report(study_id, code, study_date, modality,"
        " slices, report, status, folder) VALUES (?,?,?,?,?,?,?,?)",
        [(r["study_id"], fmt(r["code"], widths) or None, r["date"],
          r["raw_modality"], r["slices"], r["report"], r["status"], r["dir"])
         for r in out])
    conn.commit()


def suggest(studies, reports):
    """Propose codes for studies whose folder carries none.

    Reports are named with code + date + modality; DICOM headers carry date and
    modality. That overlap is the only bridge back to a code for a study in an
    unnamed folder. A lead, not a match - hence its own sheet.
    """
    index = defaultdict(set)
    for code, items in reports.items():
        for _, date, modality, _ in items:
            if date:
                index[(date, modality)].add(code)
                index[(date, None)].add(code)

    out = []
    for s in studies:
        if s["status"] != "unassigned - no code in path":
            continue
        cands = sorted(index.get((s["date"], s["raw_modality"] or None))
                       or index.get((s["date"], None)) or set())
        if len(cands) == 1:
            conf, note = "strong", "only patient with a report on this date+modality"
        elif 2 <= len(cands) <= 5:
            conf, note = "weak", f"{len(cands)} patients share this date+modality"
        elif cands:
            conf, note = "none", f"{len(cands)} candidates - too many to be useful"
            cands = cands[:5] + [("...", 0)]
        else:
            conf, note = "none", "no report anywhere matches this date+modality"
        out.append(dict(dir=s["dir"], date=s["date"], modality=s["modality"],
                        slices=s["slices"], candidates=cands,
                        confidence=conf, note=note))
    return out


# --------------------------------------------------------------------------
# Excel output
# --------------------------------------------------------------------------

from openpyxl import Workbook                             # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter              # noqa: E402

HEADER_FILL = PatternFill("solid", fgColor="DDDDDD")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")


def sheet(wb, title, header, widths, rows):
    ws = wb.create_sheet(title)
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in rows:
        ws.append(row)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def human_bytes(n):
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:,} B" if unit == "B" else f"{n:,.1f} {unit}"
        n /= 1024


def write_workbook(data, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    W = data["widths"]

    # Per-patient verdicts from this tool's own matching, so the Coverage
    # sheet and the Studies sheet can never disagree.
    cov_detail = coverage.new_detail()
    for s_ in data["studies"]:
        if not s_["code"]:
            continue
        code = fmt(s_["code"], W)
        if s_["status"] == "report with no matching images":
            cov_detail[code][coverage.ORPHAN] += 1
        elif s_["status"].startswith("matched"):
            cov_detail[code][coverage.MATCHED] += 1
        else:
            cov_detail[code][coverage.UNMATCHED] += 1
    cov_rows, cov_summary, cov_totals = coverage.classify(data["conn"], cov_detail)

    studies = data["studies"]
    real = [s for s in studies if s["study_id"] is not None]
    assigned = [s for s in real if s["code"]]
    with_report = [s for s in assigned if s["status"].startswith("matched")]
    patients = sorted((c for c in data["per_patient"] if c != "UNASSIGNED"),
                      key=lambda c: (c[0], c[1]))
    studies_by_code = Counter(s["code"] for s in assigned)
    reported_by_code = Counter(s["code"] for s in with_report)

    ov = [
        ("Archive root", data["root"]),
        ("Study prefixes found", ", ".join(sorted(W))),
        ("", ""),
        ("Patients found", len(patients)),
        ("Total files", sum(data["kind_totals"].values())),
        ("Total size", human_bytes(sum(data["per_patient_bytes"].values()))),
        ("", ""),
        ("DICOM studies", len(real)),
        ("  attributed to a patient", len(assigned)),
        ("  NOT attributed (see Unassigned)", len(real) - len(assigned)),
        ("", ""),
        ("Studies with a matching report", len(with_report)),
        ("Studies without a report", len(assigned) - len(with_report)),
        ("Reports with no matching images",
         sum(1 for s in studies if s["status"] == "report with no matching images")),
        ("", ""),
        ("Patients with at least one study", len(studies_by_code)),
        ("Patients where every study has a report",
         sum(1 for c in studies_by_code if reported_by_code[c] == studies_by_code[c])),
        ("", ""),
    ] + coverage.summary_lines(cov_summary, cov_totals) + [
        ("", ""),
        ("Filename/folder code conflicts", len(data["conflicts"])),
        ("PDFs with no patient code at all", len(data["orphan_reports"])),
        ("", ""),
    ] + [(f"Files of kind '{k}'", v) for k, v in data["kind_totals"].most_common()]
    ws = sheet(wb, "Overview", ["Measure", "Value"], (48, 62), ov)
    ws.auto_filter.ref = None

    header = (["Patient Code"] + [k.upper() for k in KIND_COLUMNS]
              + ["Total Files", "Total Size", "Studies", "Studies w/ Report",
                 "Studies w/o Report"])
    widths = [14] + [9] * len(KIND_COLUMNS) + [12, 13, 9, 17, 18]
    rows = []
    for code in patients + (["UNASSIGNED"] if "UNASSIGNED" in data["per_patient"] else []):
        counts = data["per_patient"][code]
        label = code if code == "UNASSIGNED" else fmt(code, W)
        n_stud = studies_by_code.get(code, 0)
        n_rep = reported_by_code.get(code, 0)
        rows.append([label] + [counts.get(k, 0) for k in KIND_COLUMNS]
                    + [sum(counts.values()),
                       human_bytes(data["per_patient_bytes"][code]),
                       n_stud, n_rep, n_stud - n_rep])
    ws = sheet(wb, "Patients", header, widths, rows)
    for row in ws.iter_rows(min_row=2):
        if row[0].value == "UNASSIGNED" or (row[-1].value or 0) > 0:
            for cell in row:
                cell.fill = WARN_FILL

    sheet(wb, "Studies",
          ["Patient Code", "Study Date", "Modality", "DICOM Files",
           "Count Confidence", "Matched Report", "Status", "Folder"],
          (14, 12, 12, 12, 16, 42, 38, 70),
          [[fmt(s["code"], W), s["date"] or "", s["modality"], s["slices"],
            s["confidence"], s["report"], s["status"], s["dir"]]
           for s in sorted(studies, key=lambda s: (fmt(s["code"], W), s["date"] or ""))])

    sheet(wb, "Unassigned",
          ["Folder", "Files With No Code", "Note"], (95, 18, 40),
          [[data["dir_paths"][did], n, "no patient code anywhere in this path"]
           for did, n in sorted(data["unassigned_dirs"].items(), key=lambda kv: -kv[1])])

    sheet(wb, "Coverage", coverage.HEADER, coverage.WIDTHS, cov_rows)

    sheet(wb, "Suggestions",
          ["Folder", "Study Date", "Modality", "DICOM Files",
           "Candidate Code(s)", "Confidence", "Why"],
          (70, 12, 12, 12, 26, 12, 46),
          [[s["dir"], s["date"] or "", s["modality"], s["slices"],
            ", ".join(c if isinstance(c, str) else fmt(c, W) for c in s["candidates"]),
            s["confidence"], s["note"]]
           for s in sorted(data["suggestions"],
                           key=lambda s: (s["confidence"] != "strong", s["dir"]))])

    sheet(wb, "Conflicts",
          ["Assigned Code", "Other Code Seen", "Assigned From", "File"],
          (15, 16, 20, 95),
          [[fmt(c[0], W), fmt(c[1], W), c[2], c[3]] for c in data["conflicts"]])

    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--out", default="master_report.xlsx")
    ap.add_argument("--discover", action="store_true")
    ap.add_argument("--prefixes", default=None)
    ap.add_argument("--code-regex", default=None)
    ap.add_argument("--date-order", default="dmy", choices=["dmy", "mdy"])
    ap.add_argument("--pad", default=None)
    ap.add_argument("--prefer", default="filename", choices=["filename", "folder"])
    ap.add_argument("--csv", default=None)
    args = ap.parse_args()

    if not HAS_DATE_ORDER:
        print("note: your match_reports.py is an older copy whose "
              "parse_pdf_date() takes no date_order.", file=sys.stderr)
        if args.date_order != "dmy":
            print(f"      --date-order {args.date_order} CANNOT be applied. "
                  "Copy the newer match_reports.py if your PDF dates are "
                  "month-day-year.", file=sys.stderr)

    if args.discover:
        discover(sqlite3.connect(args.db))
        return

    data = build(args)
    write_workbook(data, args.out)

    real = [s for s in data["studies"] if s["study_id"] is not None]
    matched = sum(1 for s in real if s["status"].startswith("matched"))
    unassigned = sum(1 for s in real if not s["code"])
    strong = sum(1 for s in data["suggestions"] if s["confidence"] == "strong")
    print(f"wrote {args.out}")
    print(f"  prefixes: {', '.join(sorted(data['widths']))}")
    print(f"  patients: {len([c for c in data['per_patient'] if c != 'UNASSIGNED']):,}")
    print(f"  studies:  {len(real):,}  ({matched:,} with a report, "
          f"{unassigned:,} unattributed)")
    print(f"  conflicts: {len(data['conflicts']):,}   "
          f"strong suggestions: {strong:,}")
    print(f"  codes written to the database - query the v_files view")
    if args.csv:
        print(f"  file-level inventory: {args.csv}")


if __name__ == "__main__":
    main()
