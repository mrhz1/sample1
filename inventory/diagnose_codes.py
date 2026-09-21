"""Why did my codes come out wrong? - read-only inspection of inventory.db.

Answers the three questions that come up when attribution looks off:

  1. Is this database actually attributed, and by what?
  2. Why is every code padded wider than the real ones (AA0001 -> AA000001)?
  3. Why did report.py --discover not list a prefix that clearly exists?

Writes nothing. Safe to run against a live database mid-pipeline.

Usage:
    python diagnose_codes.py --db inventory.db --prefixes AA

    --db PATH        The database to inspect.
    --prefixes LIST  Prefixes to analyse, as passed to report.py. Defaults to
                     whatever is already stored in the files table.
    --code-regex RE  Analyse this pattern instead, as passed to report.py.
    --examples N     Example names to show per bucket (default 5).
"""

import argparse
import os
import re
import sqlite3
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from report import (  # noqa: E402
    CANDIDATE_RE,
    build_code_re,
    parse_code,
    pattern_hints,
    resolve_widths,
)

RULE = "-" * 78


def head(title):
    print(f"\n{title}\n{RULE}")


def scalar(conn, sql, *args):
    row = conn.execute(sql, args).fetchone()
    return row[0] if row else None


def section_database(conn, db_path):
    head(f"1. DATABASE  {db_path}")
    size = os.path.getsize(db_path) / 1e6
    print(f"  {size:,.1f} MB on disk")
    for key, value in conn.execute("SELECT key, value FROM meta ORDER BY key"):
        print(f"  meta.{key:<16} {value}")
    for table, label in [("dirs", "directories"), ("files", "files"),
                         ("studies", "studies"), ("dicom_meta", "DICOM headers")]:
        try:
            print(f"  {scalar(conn, f'SELECT COUNT(*) FROM {table}'):>12,}  {label}")
        except sqlite3.OperationalError:
            print(f"  {'-':>12}  {label} (table missing - that pass has not run)")


def section_attribution(conn, examples):
    """What report.py actually wrote. If this section is empty, nothing else
    in the pipeline can produce a code - including the Excel output."""
    head("2. ATTRIBUTION AS STORED")
    n_files = scalar(conn, "SELECT COUNT(*) FROM files")
    n_coded = scalar(conn, "SELECT COUNT(*) FROM files WHERE code IS NOT NULL")
    n_prefix = scalar(conn, "SELECT COUNT(*) FROM files WHERE prefix IS NOT NULL")
    n_dirs = scalar(conn, "SELECT COUNT(*) FROM dirs WHERE code IS NOT NULL")
    print(f"  files with a code:        {n_coded:>12,} of {n_files:,}")
    print(f"  files with a prefix:      {n_prefix:>12,}")
    print(f"  directories with a code:  {n_dirs:>12,}")

    if not n_coded and not n_prefix:
        print("\n  NOTHING IS ATTRIBUTED IN THIS DATABASE.")
        print("  If an Excel file shows codes, it was built from a different")
        print("  database or an earlier run - report.py clears every code at the")
        print("  start of each run, so the last run is all that survives here.")
        return []

    prefixes = []
    print(f"\n  {'PREFIX':<10} {'FILES':>10} {'PATIENTS':>9} {'NUMBERS':>13}"
          f"  STORED CODE WIDTHS")
    for prefix, n, lo, hi in conn.execute(
            "SELECT prefix, COUNT(*), MIN(number), MAX(number) FROM files"
            " WHERE prefix IS NOT NULL GROUP BY prefix ORDER BY COUNT(*) DESC"):
        prefixes.append(prefix)
        patients = scalar(
            conn, "SELECT COUNT(DISTINCT number) FROM files WHERE prefix = ?",
            prefix)
        widths = Counter(
            len(code) - len(prefix) for (code,) in conn.execute(
                "SELECT DISTINCT code FROM files WHERE prefix = ?"
                " AND code IS NOT NULL", (prefix,)))
        w = ", ".join(f"{k} digits({v:,})" for k, v in sorted(widths.items()))
        print(f"  {prefix:<10} {n:>10,} {patients:>9,} {lo:>6,}-{hi:<6,}  {w}")

    print("\n  example stored codes:")
    for (code, path, name) in conn.execute(
            "SELECT f.code, d.path, f.name FROM files f JOIN dirs d"
            " ON d.id = f.dir_id WHERE f.code IS NOT NULL"
            " GROUP BY f.code LIMIT ?", (examples,)):
        print(f"    {code:<14} {os.path.join(path, name)[-70:]}")
    return prefixes


def iter_names(conn, hints):
    """Every directory-name part and file name that could hold a code.

    Filtered in SQL first - on a multi-million-row archive, running the regex
    over every file name in Python is the slow way to ask this.
    """
    upper_hints = [h.upper() for h in hints]
    seen_dirs = set()
    for (path,) in conn.execute("SELECT path FROM dirs"):
        for part in path.replace("\\", "/").split("/"):
            if not part or part in seen_dirs:
                continue
            seen_dirs.add(part)
            if any(h in part.upper() for h in upper_hints):
                yield "folder", part
    for hint in hints:
        for (name,) in conn.execute(
                "SELECT DISTINCT name FROM files WHERE name LIKE ?",
                (f"%{hint}%",)):
            yield "file", name


def section_padding(conn, code_re, hints, examples):
    """Codes are padded to the most common digit width seen for that prefix.
    Any other width is either a real variant or a name that is not a code."""
    head("3. DIGIT WIDTHS AND PADDING")
    print(f"  pattern: {code_re.pattern}")
    buckets = defaultdict(lambda: defaultdict(list))
    counts = defaultdict(Counter)
    for kind, name in iter_names(conn, hints):
        found = parse_code(name, code_re)
        if not found:
            continue
        prefix, _number, width = found
        counts[prefix][width] += 1
        if len(buckets[prefix][width]) < examples:
            buckets[prefix][width].append(f"{kind}: {name}")

    if not counts:
        print("  this pattern matches NOTHING in the archive.")
        return
    chosen = resolve_widths(counts)
    for prefix in sorted(counts, key=lambda p: -sum(counts[p].values())):
        width = chosen[prefix]
        print(f"\n  {prefix}: padded to {width} digits"
              f"  ->  {prefix}{1:0{width}d}")
        for w in sorted(counts[prefix]):
            n = counts[prefix][w]
            flag = "  <-- used" if w == width else ""
            print(f"    {w} digits  {n:>9,} name(s){flag}")
            for ex in buckets[prefix][w]:
                print(f"        {ex[:68]}")
        odd = {w: n for w, n in counts[prefix].items() if w != width and n < 20}
        if odd:
            print(f"    ^ the {sum(odd.values()):,} name(s) at other widths are"
                  " worth a look - a date or an\n      accession number caught by"
                  " the pattern shows up exactly like this.")
            print("      They no longer affect padding, but they are still"
                  " attributed as patients.")


def section_discover(conn, hints, examples):
    """--discover uses a word-bounded pattern; the --prefixes pattern is not
    bounded at all. Names where they disagree are why a prefix goes missing."""
    head("4. WHY --discover MAY NOT LIST YOUR PREFIX")
    print(f"  --discover pattern: {CANDIDATE_RE.pattern}")
    print("  (word-bounded on both sides - the codes must stand alone)\n")
    yields = defaultdict(list)
    counts = Counter()
    for _kind, name in iter_names(conn, hints):
        found = CANDIDATE_RE.findall(name)
        if found:
            for prefix, digits in found:
                key = f"reported as {prefix.upper()}"
                counts[key] += 1
                if len(yields[key]) < examples:
                    yields[key].append(name)
        else:
            counts["NOT SEEN BY --discover"] += 1
            if len(yields["NOT SEEN BY --discover"]) < examples:
                yields["NOT SEEN BY --discover"].append(name)
    if not counts:
        print("  no candidate names found at all.")
        return
    for key, n in counts.most_common():
        print(f"  {key:<34} {n:>9,} name(s)")
        for ex in yields[key]:
            print(f"      {ex[:68]}")
    if "NOT SEEN BY --discover" in counts:
        print("\n  Names in that last group glue the code to other characters, so"
              "\n  the word-bounded pattern skips them entirely and the prefix can"
              "\n  be missing from the --discover table while --prefixes, which is"
              "\n  not bounded, still matches thousands of files.")


def section_rejected(conn, code_re, prefixes, digits, hints, examples):
    """Names a looser pattern would accept but this one rejects, by reason.

    Tightening a rule is invisible until something disappears, and by then the
    patient count has moved and nobody knows which rule took them. This puts a
    number against each guard before you trust the run.
    """
    head("5. NAMES THIS PATTERN REJECTS")
    loose = build_code_re(prefixes, None, "1-9")
    loose = re.compile(loose.pattern.replace("(?<![A-Za-z0-9])", "")
                       .replace("(?![A-Za-z0-9])", ""), re.IGNORECASE)
    lo, hi = None, None
    try:
        from report import parse_digits
        lo, hi = parse_digits(digits)
    except Exception:
        lo, hi = 1, 5

    reasons = defaultdict(list)
    counts = Counter()
    for kind, name in iter_names(conn, hints):
        if parse_code(name, code_re):
            continue
        m = loose.search(name)
        if not m:
            continue
        start, end = m.span()
        run = re.search(r"\d+$", m.group(0))
        width = len(run.group(0)) if run else 0
        if start > 0 and name[start - 1].isalnum():
            why = "a letter or digit runs into the prefix"
        elif end < len(name) and name[end].isalnum():
            why = "a letter or digit follows the number"
        elif not (lo <= width <= hi):
            why = f"{width} digits, outside --digits {digits or '1-5'}"
        else:
            why = "other"
        counts[why] += 1
        if len(reasons[why]) < examples:
            reasons[why].append(f"{kind}: {name}")

    if not counts:
        print("  nothing is rejected - every code-like name is being attributed.")
        return
    print("  These would be codes under a looser rule. Each line is a patient")
    print("  you do NOT have, so check the examples before accepting them.\n")
    for why, n in counts.most_common():
        print(f"  {n:>9,}  {why}")
        for ex in reasons[why]:
            print(f"             {ex[:64]}")


# The guards, innermost first, exactly as build_code_re layers them. Each row
# is (label, leading guard, trailing guard).
VARIANTS = [
    ("no guards (the original rule)", "", ""),
    ("+ no longer digit run", "", r"(?!\d)"),
    ("+ no letter after the number", "", r"(?![A-Za-z0-9])"),
    ("+ no letter before the prefix", r"(?<![A-Za-z0-9])", r"(?![A-Za-z0-9])"),
]


def section_compare(conn, prefixes, digits, hints, examples):
    """Patients found under each rule, so the cost of each guard is a number.

    Run this instead of re-running report.py with one flag changed at a time:
    it scans the same names once and reports what every variant would have
    attributed, plus what the tighter rule dropped.
    """
    head("6. WHAT EACH RULE WOULD ATTRIBUTE")
    alt = "|".join(re.escape(p.strip()) for p in prefixes.split(",") if p.strip())
    lo, hi = 1, 5
    names = list(iter_names(conn, hints))

    results = []
    for label, lead, trail in VARIANTS:
        rx = re.compile(rf"{lead}(?:{alt})[-_ ]?\d{{{lo},{hi}}}{trail}",
                        re.IGNORECASE)
        found = {}
        for kind, name in names:
            got = parse_code(name, rx)
            if got:
                found.setdefault((got[0], got[1]), f"{kind}: {name}")
        results.append((label, found))

    if digits:
        from report import parse_digits
        dlo, dhi = parse_digits(digits)
        lead, trail = VARIANTS[-1][1], VARIANTS[-1][2]
        rx = re.compile(rf"{lead}(?:{alt})[-_ ]?\d{{{dlo},{dhi}}}{trail}",
                        re.IGNORECASE)
        found = {}
        for kind, name in names:
            got = parse_code(name, rx)
            if got:
                found.setdefault((got[0], got[1]), f"{kind}: {name}")
        results.append((f"+ --digits {digits}", found))

    print(f"  {'RULE':<34} {'PATIENTS':>9}  {'LOST HERE':>9}")
    print("  " + "-" * 74)
    previous = None
    for label, found in results:
        lost = 0 if previous is None else len(set(previous) - set(found))
        print(f"  {label:<34} {len(found):>9,}  {lost:>9,}")
        if previous is not None and lost:
            for code in list(set(previous) - set(found))[:examples]:
                print(f"      dropped {code[0]}{code[1]}: {previous[code][:52]}")
        previous = found
    print("\n  The last line is what report.py is using now. If the patient")
    print("  count you expect is on an earlier line, that guard is the one")
    print("  costing you - pass --code-regex to reproduce it exactly:")
    print(f"      --code-regex \"(?:{alt})[-_ ]?[0-9]{{1,5}}\"")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--db", default="inventory.db")
    ap.add_argument("--prefixes", default=None)
    ap.add_argument("--code-regex", default=None)
    ap.add_argument("--digits", default="1-5")
    ap.add_argument("--guards", default="both")
    ap.add_argument("--examples", type=int, default=5)
    args = ap.parse_args()

    if not os.path.exists(args.db):
        sys.exit(f"no such database: {args.db}")
    conn = sqlite3.connect(f"file:{os.path.abspath(args.db)}?mode=ro", uri=True)

    section_database(conn, args.db)
    stored = section_attribution(conn, args.examples)

    prefixes = args.prefixes or ",".join(stored)
    if not prefixes and not args.code_regex:
        print("\nGive --prefixes to analyse padding and --discover behaviour.")
        return
    code_re = build_code_re(prefixes, args.code_regex, args.digits,
                            args.guards)
    hints = sorted(set(pattern_hints(code_re.pattern)))

    section_padding(conn, code_re, hints, args.examples)
    section_discover(conn, hints, args.examples)
    section_rejected(conn, code_re, prefixes, args.digits, hints, args.examples)
    if prefixes:
        section_compare(conn, prefixes, args.digits, hints, args.examples)
    print()


if __name__ == "__main__":
    main()
