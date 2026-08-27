"""List every PDF found under a folder, straight into the log.

No Excel, no copying, no renaming - it just walks the folder you point it at
and logs the full path of every PDF it finds, then a count at the end.

Usage:
    python list_pdfs.py <folder> [log_file] [--no-recursive]

    folder          Folder to search.
    log_file        Where to append the listing (default: <folder>/pdf_list.log).
    --no-recursive  Only the folder itself, don't descend into subfolders.

Examples:
    python list_pdfs.py "/mnt/data/reports"
    python list_pdfs.py "/mnt/data/reports" pdfs.log
    python list_pdfs.py "/mnt/data/reports" --no-recursive
"""

import sys
from datetime import datetime
from pathlib import Path

_log_fh = None


def open_log(path):
    """Append-mode log so successive runs stack up in one place."""
    global _log_fh
    _log_fh = open(path, "a", encoding="utf-8")
    _log_fh.write(f"\n===== run started {datetime.now():%Y-%m-%d %H:%M:%S} =====\n")
    _log_fh.flush()


def log(message=""):
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {message}" if message else ""
    print(line, flush=True)
    if _log_fh is not None:
        _log_fh.write(line + "\n")
        _log_fh.flush()


def human_size(num_bytes):
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024


def find_pdfs(folder, recursive=True):
    """Every *.pdf under the folder, case-insensitive, sorted by path.

    Skips anything unreadable rather than dying halfway through a big tree.
    """
    pattern = "**/*" if recursive else "*"
    found = []
    try:
        entries = folder.glob(pattern)
    except OSError as exc:
        log(f"  WARNING: could not read {folder}: {exc}")
        return found
    while True:
        try:
            path = next(entries)
        except StopIteration:
            break
        except OSError as exc:
            log(f"  WARNING: skipped an unreadable entry: {exc}")
            continue
        if path.suffix.lower() == ".pdf" and path.is_file():
            found.append(path)
    return sorted(found)


def parse_args(argv):
    recursive = "--no-recursive" not in argv
    positional = [a for a in argv if not a.startswith("--")]
    if not positional:
        print(__doc__)
        return None
    folder = Path(positional[0]).expanduser()
    log_file = Path(positional[1]).expanduser() if len(positional) > 1 else None
    return folder, log_file, recursive


def main(argv):
    if argv and argv[0] in ("-h", "--help"):
        print(__doc__)
        return 0

    parsed = parse_args(argv)
    if parsed is None:
        return 1
    folder, log_file, recursive = parsed

    if not folder.is_dir():
        print(f"Folder not found: {folder}")
        return 1

    if log_file is None:
        log_file = folder / "pdf_list.log"
    log_file.parent.mkdir(parents=True, exist_ok=True)
    open_log(log_file)

    log(f"Folder:    {folder.resolve()}")
    log(f"Recursive: {'yes' if recursive else 'no (this folder only)'}")
    log(f"Log file:  {log_file}")
    log()

    pdfs = find_pdfs(folder, recursive)
    if not pdfs:
        log("No PDF files found.")
        return 0

    total_bytes = 0
    for index, path in enumerate(pdfs, 1):
        try:
            size = path.stat().st_size
        except OSError:
            size = 0
        total_bytes += size
        log(f"  {index:>5}. {path.resolve()}  ({human_size(size)})")

    log()
    log(f"Total PDFs found: {len(pdfs):,}  ({human_size(total_bytes)})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
