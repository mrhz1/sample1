"""How many files are under this path, and how much do they add up to?

The two numbers Windows shows in right-click -> Properties, for paths where
right-clicking is not an option: a network share over a slow link, a folder
with millions of files, or a machine you only have a shell on.

Sizes are reported the way Windows reports them - KB/MB/GB/TB in multiples of
1024, with the exact byte count in brackets - so the output can be compared
against the Properties dialog without having to convert anything.

Usage:
    python folder_size.py <path> [<path> ...]

    --quiet     No progress line while it runs.
    --bytes     Print the raw byte count only, for scripting.

Counts files, not directories. A link to a folder - a symlink or a Windows
junction - is never followed and is not counted as a file, so a share that
links back into itself is neither counted twice nor able to hang the scan.
"""

import argparse
import os
import sys
import time

# Mirrors crawl.py's long_path(). Inlined rather than imported so this file
# stays a single-file tool you can copy onto a machine by itself - without it,
# a path over 260 characters silently reads as an empty directory on Windows,
# and the count comes back wrong rather than failing.
def long_path(path):
    if os.name != "nt":
        return path
    p = os.path.abspath(path)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def human(n):
    """Windows-style: 1024-based, two decimals, same unit names it uses."""
    size = float(n)
    for unit in ("bytes", "KB", "MB", "GB", "TB", "PB"):
        if size < 1024 or unit == "PB":
            if unit == "bytes":
                return f"{int(size)} bytes"
            return f"{size:,.2f} {unit}"
        size /= 1024


def walk(root, on_progress=None):
    """(files, bytes, errors), counted with one directory read per directory.

    os.scandir carries size in the directory entry on Windows, so this never
    stats a file individually - the difference between one round trip per
    directory and one per file, which is what makes a million-file share
    finish in minutes instead of hours.
    """
    files = total = errors = 0
    stack = [root]
    while stack:
        current = stack.pop()
        try:
            with os.scandir(long_path(current)) as it:
                for entry in it:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            stack.append(entry.path)
                        elif entry.is_symlink() and entry.is_dir():
                            # A link to a folder is neither a file nor a folder
                            # to descend into. Windows junctions land here too,
                            # which is what stops a share that links back into
                            # itself being counted twice.
                            continue
                        else:
                            files += 1
                            total += entry.stat(follow_symlinks=False).st_size
                    except OSError:
                        errors += 1
        except OSError:
            errors += 1
        if on_progress:
            on_progress(files, total, len(stack))
    return files, total, errors


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--quiet", action="store_true")
    ap.add_argument("--bytes", action="store_true",
                    help="print the byte count only")
    args = ap.parse_args()

    status = {"last": 0.0}

    def progress(files, total, pending):
        if args.quiet or args.bytes or not sys.stderr.isatty():
            return
        now = time.time()
        if now - status["last"] < 0.5:
            return
        status["last"] = now
        print(f"\r  {files:,} files, {human(total)}, {pending:,} folders left"
              "   ", end="", file=sys.stderr, flush=True)

    exit_code = 0
    for path in args.paths:
        if not os.path.isdir(long_path(path)):
            print(f"not a directory: {path}", file=sys.stderr)
            exit_code = 1
            continue

        started = time.time()
        files, total, errors = walk(path, progress)
        if not args.quiet and not args.bytes and sys.stderr.isatty():
            print("\r" + " " * 70 + "\r", end="", file=sys.stderr)

        if args.bytes:
            print(total)
            continue
        print(path)
        print(f"  Files:  {files:,}")
        print(f"  Size:   {human(total)} ({total:,} bytes)")
        if errors:
            print(f"  Skipped {errors:,} item(s) that could not be read"
                  " (permissions, or a path that vanished mid-scan)")
            exit_code = 1
        elapsed = time.time() - started
        if elapsed > 5:
            print(f"  ({elapsed / 60:.1f} min)")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
