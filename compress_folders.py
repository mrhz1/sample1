#!/usr/bin/env python3
"""
Compress a range of exam-code folders into individual Linux tar archives,
processing many folders in parallel with multiple worker processes.

Layout:
    <root>/AA001/...
    <root>/AA002/...
    ...
    <root>/AA1400/...

Folder numbering is zero-padded to at least the width of --start's number
but grows beyond that without truncation (AA001 ... AA999, AA1000 ...
AA1400), matching the naming scheme in use.

Each folder becomes exactly one archive file (e.g. AA001.tar.gz) written to
the output directory, with the folder's contents stored under an arcname of
the code itself (so extracting the archive recreates "AA001/...").

Usage:
    python compress_folders.py --root "/mnt/data" --start AA001 --end AA1400 \
        --output-dir "/mnt/archives" --workers 8

    # Plain (uncompressed) POSIX tar instead of gzip:
    python compress_folders.py --root "/mnt/data" --start AA001 --end AA1400 \
        --output-dir "/mnt/archives" --compression none

Re-running the same command skips codes whose archive file already exists
unless --force is given.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import logging
import re
import sys
import tarfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

CODE_PATTERN = re.compile(r"^(?P<prefix>[A-Za-z]+)(?P<num>\d+)$")

RETRY_ATTEMPTS = 3
RETRY_DELAY_SECONDS = 2

COMPRESSION_EXT = {"gz": ".tar.gz", "bz2": ".tar.bz2", "xz": ".tar.xz", "none": ".tar"}
COMPRESSION_MODE = {"gz": "w:gz", "bz2": "w:bz2", "xz": "w:xz", "none": "w"}

logger = logging.getLogger("compress_folders")


@dataclass
class CompressResult:
    code: str
    status: str  # "done", "skipped", or "error"
    archive_path: str = ""
    size_bytes: int = 0
    error: str = ""


def retry(func, *args, **kwargs):
    """Retry a filesystem operation a few times to ride out transient
    network hiccups before giving up."""
    last_exc = None
    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            return func(*args, **kwargs)
        except OSError as exc:
            last_exc = exc
            logger.warning(
                "IO error (attempt %d/%d) on %s: %s", attempt, RETRY_ATTEMPTS, args, exc
            )
            time.sleep(RETRY_DELAY_SECONDS)
    raise last_exc


def parse_code(code: str) -> tuple[str, str, int]:
    """Split a code like 'AA0001' into (prefix, numeric_str, width)."""
    m = CODE_PATTERN.match(code)
    if not m:
        raise ValueError(f"'{code}' does not look like a code (e.g. AA0001)")
    return m.group("prefix").upper(), m.group("num"), len(m.group("num"))


def generate_codes(start: str, end: str) -> list[str]:
    """Enumerate codes from start to end inclusive, e.g. AA001..AA1400.

    Numeric part is zero-padded to at least the width of --start's number,
    but grows beyond that without truncation (e.g. AA001, AA002, ... AA999,
    AA1000, AA1001, ... AA1400) since that's the padding scheme in use."""
    start_prefix, start_num, width = parse_code(start)
    end_prefix, end_num, _ = parse_code(end)

    if start_prefix != end_prefix:
        raise ValueError(f"--start and --end must share a prefix ('{start_prefix}' vs '{end_prefix}')")

    lo, hi = int(start_num), int(end_num)
    if lo > hi:
        raise ValueError(f"--start ({start}) is after --end ({end})")

    return [f"{start_prefix}{n:0{width}d}" for n in range(lo, hi + 1)]


def compress_folder(code: str, src_dir: str, output_dir: str, compression: str, force: bool) -> CompressResult:
    src = Path(src_dir)
    ext = COMPRESSION_EXT[compression]
    mode = COMPRESSION_MODE[compression]
    archive_path = Path(output_dir) / f"{code}{ext}"

    if archive_path.exists() and not force:
        return CompressResult(code=code, status="skipped", archive_path=str(archive_path))

    tmp_path = archive_path.with_suffix(archive_path.suffix + ".part")
    try:
        def _write():
            with tarfile.open(tmp_path, mode) as tar:
                tar.add(src, arcname=code)

        retry(_write)
        tmp_path.replace(archive_path)
        size = archive_path.stat().st_size
        return CompressResult(code=code, status="done", archive_path=str(archive_path), size_bytes=size)
    except Exception as exc:  # noqa: BLE001 - a folder-level failure must not kill the run
        tmp_path.unlink(missing_ok=True)
        return CompressResult(code=code, status="error", error=f"{type(exc).__name__}: {exc}")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--root", required=True, help="Parent directory containing the code folders")
    parser.add_argument("--start", required=True, help="First code, e.g. AA0001")
    parser.add_argument("--end", required=True, help="Last code (inclusive), e.g. AA0100")
    parser.add_argument("--output-dir", required=True, help="Directory to write archive files into")
    parser.add_argument("--workers", type=int, default=4, help="Number of parallel worker processes")
    parser.add_argument(
        "--compression",
        choices=sorted(COMPRESSION_EXT),
        default="gz",
        help="Compression to use (default: gz). 'none' produces a plain POSIX .tar",
    )
    parser.add_argument("--force", action="store_true", help="Recompress even if the archive already exists")
    parser.add_argument("--log-file", default=None, help="Optional log file path")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler()] + ([logging.FileHandler(args.log_file)] if args.log_file else []),
    )

    root = Path(args.root)
    if not root.is_dir():
        parser.error(f"--root is not a directory: {root}")

    output_dir = Path(args.output_dir).resolve()
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        probe = output_dir / ".compress_folders_write_test"
        probe.write_bytes(b"")
        probe.unlink()
    except OSError as exc:
        parser.error(f"--output-dir is not writable: {output_dir} ({exc})")

    try:
        codes = generate_codes(args.start, args.end)
    except ValueError as exc:
        parser.error(str(exc))
        return

    todo: list[tuple[str, Path]] = []
    missing: list[str] = []
    for code in codes:
        folder = root / code
        if folder.is_dir():
            todo.append((code, folder))
        else:
            missing.append(code)

    if missing:
        logger.warning("%d codes not found under %s (skipped): %s", len(missing), root, missing)

    logger.info(
        "%d codes total, %d found, compressing with %d workers (compression=%s)",
        len(codes), len(todo), args.workers, args.compression,
    )

    if not todo:
        logger.info("Nothing to do.")
        return

    done = skipped = errors = 0
    total_bytes = 0
    completed = 0
    with concurrent.futures.ProcessPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(compress_folder, code, str(folder), str(output_dir), args.compression, args.force): code
            for code, folder in todo
        }
        for future in concurrent.futures.as_completed(futures):
            code = futures[future]
            try:
                result = future.result()
            except Exception as exc:  # noqa: BLE001
                result = CompressResult(code=code, status="error", error=f"worker crashed: {exc}")

            completed += 1
            if result.status == "done":
                done += 1
                total_bytes += result.size_bytes
                logger.info(
                    "[%d/%d] %s -> %s (%.1f MB)",
                    completed, len(todo), code, result.archive_path, result.size_bytes / 1_048_576,
                )
            elif result.status == "skipped":
                skipped += 1
                logger.info("[%d/%d] %s skipped (already exists)", completed, len(todo), code)
            else:
                errors += 1
                logger.error("[%d/%d] %s FAILED: %s", completed, len(todo), code, result.error)

    logger.info(
        "Finished: %d done (%.1f MB total), %d skipped, %d errors, %d not found",
        done, total_bytes / 1_048_576, skipped, errors, len(missing),
    )
    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
