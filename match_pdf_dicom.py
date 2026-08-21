#!/usr/bin/env python3
"""
Match PDF reports (Folder A) to DICOM image sets (Folder B) by patient/exam code,
modality and date, and produce an Excel report of how many DICOM files were
found for each PDF.

Folder A layout:
    <folder_a>/AA0001/.../AA0001-MRI-2017-06-16.pdf

Folder B layout:
    <folder_b>/AA0001/.../<extensionless DICOM files, possibly mixed with other files>

Usage:
    python match_pdf_dicom.py --folder-a "\\server\A" --folder-b "\\server\B" \
        --output report.xlsx --workers 8

Re-running the same command after a crash / network drop will skip codes that
already completed successfully (tracked in the checkpoint SQLite database) and
only reprocess codes that are pending or previously errored.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import logging
import re
import sqlite3
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

try:
    import pydicom
    from pydicom.errors import InvalidDicomError
except ImportError:
    print("Missing dependency 'pydicom'. Install with: pip install pydicom", file=sys.stderr)
    raise

try:
    import pandas as pd
except ImportError:
    print("Missing dependency 'pandas'. Install with: pip install pandas openpyxl", file=sys.stderr)
    raise


# --------------------------------------------------------------------------
# Config / constants
# --------------------------------------------------------------------------

CODE_DIR_PATTERN = re.compile(r"^[A-Za-z]{2}\d{4,}$")

# code - modality - date(yyyymmdd or yyyy-mm-dd), tolerant of spacing / separators
PDF_NAME_PATTERN = re.compile(
    r"^(?P<code>.+?)\s*-\s*(?P<modality>.+?)\s*-\s*(?P<date>\d{4}[-_./]?\d{2}[-_./]?\d{2})$",
    re.IGNORECASE,
)

# DICOM tags to try, in priority order, for "the" date of a DICOM instance
DICOM_DATE_TAGS = ["StudyDate", "SeriesDate", "AcquisitionDate", "ContentDate"]

RETRY_ATTEMPTS = 3
RETRY_DELAY_SECONDS = 2

logger = logging.getLogger("match_pdf_dicom")


# --------------------------------------------------------------------------
# Data classes
# --------------------------------------------------------------------------

@dataclass
class PdfRecord:
    code: str
    path: Path
    filename: str
    modality: Optional[str]
    date: Optional[str]  # normalized YYYYMMDD
    parse_ok: bool
    matched_dicom_count: int = 0
    notes: str = ""


@dataclass
class FolderResult:
    code: str
    status: str  # "done" or "error"
    error: str = ""
    pdf_records: list = field(default_factory=list)
    dicom_total: int = 0
    dicom_unmatched_no_modality: int = 0
    dicom_unmatched_ambiguous: int = 0
    non_dicom_files_skipped: int = 0


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

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


def normalize_date(raw: str) -> Optional[str]:
    """Reduce any date-ish string to YYYYMMDD, or None if it doesn't parse."""
    if not raw:
        return None
    digits = re.sub(r"\D", "", raw)
    if len(digits) != 8:
        return None
    try:
        datetime.strptime(digits, "%Y%m%d")
    except ValueError:
        return None
    return digits


def parse_pdf_filename(path: Path) -> tuple[Optional[str], Optional[str], bool]:
    """Return (modality, normalized_date, parse_ok) from a PDF filename."""
    stem = path.stem
    m = PDF_NAME_PATTERN.match(stem)
    if not m:
        return None, None, False
    modality = m.group("modality").strip().upper()
    date = normalize_date(m.group("date"))
    if not modality or not date:
        return modality or None, date, False
    return modality, date, True


def read_dicom_meta(path: Path) -> Optional[dict]:
    """Try to read a file as DICOM and pull out Modality + a usable date.
    Returns None if the file is not a readable DICOM."""
    try:
        ds = pydicom.dcmread(str(path), stop_before_pixels=True, force=True)
    except (InvalidDicomError, OSError, ValueError, Exception):  # noqa: BLE001
        return None

    modality = getattr(ds, "Modality", None)
    if not modality:
        # Not enough to call this a usable DICOM for matching purposes.
        return None
    modality = str(modality).strip().upper()

    date = None
    for tag in DICOM_DATE_TAGS:
        val = getattr(ds, tag, None)
        if val:
            date = normalize_date(str(val))
            if date:
                break

    return {"modality": modality, "date": date, "path": str(path)}


def find_code_dirs(root: Path) -> dict[str, Path]:
    """Map code (e.g. 'AA0001') -> directory path for immediate subdirs of root
    that look like exam codes."""
    result = {}
    for entry in root.iterdir():
        if entry.is_dir() and CODE_DIR_PATTERN.match(entry.name):
            result[entry.name.upper()] = entry
    return result


# --------------------------------------------------------------------------
# Per-folder worker (runs in a subprocess)
# --------------------------------------------------------------------------

def process_folder(code: str, dir_a: str, dir_b: str) -> FolderResult:
    dir_a = Path(dir_a)
    dir_b = Path(dir_b)
    try:
        pdf_paths = sorted(retry(lambda: list(dir_a.rglob("*.pdf"))))
        pdf_records: list[PdfRecord] = []
        for p in pdf_paths:
            modality, date, ok = parse_pdf_filename(p)
            note = "" if ok else "filename did not match 'code - modality - date' pattern"
            pdf_records.append(
                PdfRecord(
                    code=code,
                    path=p,
                    filename=p.name,
                    modality=modality,
                    date=date,
                    parse_ok=ok,
                    notes=note,
                )
            )

        dicom_metas = []
        non_dicom_skipped = 0
        for f in retry(lambda: list(dir_b.rglob("*"))):
            if not f.is_file():
                continue
            # Standard preamble+"DICM" files are cheap to confirm; files without
            # it (common for extensionless DICOMs) still get a real parse
            # attempt via read_dicom_meta's force=True fallback.
            meta = read_dicom_meta(f)
            if meta is None:
                non_dicom_skipped += 1
                continue
            dicom_metas.append(meta)

        unmatched_no_modality = 0
        unmatched_ambiguous = 0

        # index candidate PDFs by modality
        by_modality: dict[str, list[PdfRecord]] = {}
        for rec in pdf_records:
            if rec.modality:
                by_modality.setdefault(rec.modality, []).append(rec)

        for meta in dicom_metas:
            candidates = by_modality.get(meta["modality"])
            if not candidates:
                unmatched_no_modality += 1
                continue

            if len(candidates) == 1:
                candidates[0].matched_dicom_count += 1
                if meta["date"] and candidates[0].date and meta["date"] != candidates[0].date:
                    candidates[0].notes = (
                        candidates[0].notes + "; " if candidates[0].notes else ""
                    ) + "some matched DICOMs have a different date (matched by modality only, single candidate)"
                continue

            # multiple PDFs share this modality in the folder -> need date to disambiguate
            if not meta["date"]:
                unmatched_ambiguous += 1
                continue

            exact = [c for c in candidates if c.date == meta["date"]]
            if len(exact) == 1:
                exact[0].matched_dicom_count += 1
                continue
            if len(exact) > 1:
                # duplicate PDFs with same modality+date - shouldn't normally happen
                exact[0].matched_dicom_count += 1
                exact[0].notes = (
                    exact[0].notes + "; " if exact[0].notes else ""
                ) + "multiple PDFs share this exact modality+date; DICOM assigned to first"
                continue

            # no exact date match: fall back to nearest date among candidates that have a date
            dated_candidates = [c for c in candidates if c.date]
            if not dated_candidates:
                unmatched_ambiguous += 1
                continue
            nearest = min(
                dated_candidates,
                key=lambda c: abs(
                    datetime.strptime(c.date, "%Y%m%d") - datetime.strptime(meta["date"], "%Y%m%d")
                ),
            )
            nearest.matched_dicom_count += 1
            nearest.notes = (
                nearest.notes + "; " if nearest.notes else ""
            ) + "matched by nearest date (no exact date match among same-modality PDFs)"

        return FolderResult(
            code=code,
            status="done",
            pdf_records=pdf_records,
            dicom_total=len(dicom_metas),
            dicom_unmatched_no_modality=unmatched_no_modality,
            dicom_unmatched_ambiguous=unmatched_ambiguous,
            non_dicom_files_skipped=non_dicom_skipped,
        )

    except Exception as exc:  # noqa: BLE001 - a folder-level failure must not kill the run
        return FolderResult(code=code, status="error", error=f"{type(exc).__name__}: {exc}")


# --------------------------------------------------------------------------
# Checkpoint DB
# --------------------------------------------------------------------------

def init_checkpoint_db(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.execute(
        """CREATE TABLE IF NOT EXISTS folders (
            code TEXT PRIMARY KEY,
            status TEXT NOT NULL,
            error TEXT,
            dicom_total INTEGER,
            dicom_unmatched_no_modality INTEGER,
            dicom_unmatched_ambiguous INTEGER,
            non_dicom_files_skipped INTEGER,
            processed_at TEXT
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS pdf_results (
            code TEXT,
            pdf_filename TEXT,
            pdf_path TEXT,
            modality TEXT,
            pdf_date TEXT,
            matched_dicom_count INTEGER,
            parse_ok INTEGER,
            notes TEXT
        )"""
    )
    conn.commit()
    return conn


def save_folder_result(conn: sqlite3.Connection, result: FolderResult) -> None:
    conn.execute("DELETE FROM pdf_results WHERE code = ?", (result.code,))
    conn.execute(
        """INSERT INTO folders (code, status, error, dicom_total, dicom_unmatched_no_modality,
                                 dicom_unmatched_ambiguous, non_dicom_files_skipped, processed_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)
           ON CONFLICT(code) DO UPDATE SET
             status=excluded.status, error=excluded.error, dicom_total=excluded.dicom_total,
             dicom_unmatched_no_modality=excluded.dicom_unmatched_no_modality,
             dicom_unmatched_ambiguous=excluded.dicom_unmatched_ambiguous,
             non_dicom_files_skipped=excluded.non_dicom_files_skipped,
             processed_at=excluded.processed_at""",
        (
            result.code,
            result.status,
            result.error,
            result.dicom_total,
            result.dicom_unmatched_no_modality,
            result.dicom_unmatched_ambiguous,
            result.non_dicom_files_skipped,
            datetime.now().isoformat(timespec="seconds"),
        ),
    )
    for rec in result.pdf_records:
        conn.execute(
            """INSERT INTO pdf_results (code, pdf_filename, pdf_path, modality, pdf_date,
                                          matched_dicom_count, parse_ok, notes)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                rec.code,
                rec.filename,
                str(rec.path),
                rec.modality,
                rec.date,
                rec.matched_dicom_count,
                1 if rec.parse_ok else 0,
                rec.notes,
            ),
        )
    conn.commit()


def get_done_codes(conn: sqlite3.Connection) -> set[str]:
    rows = conn.execute("SELECT code FROM folders WHERE status = 'done'").fetchall()
    return {r[0] for r in rows}


# --------------------------------------------------------------------------
# Excel export
# --------------------------------------------------------------------------

def export_excel(conn: sqlite3.Connection, output_path: Path) -> None:
    pdf_df = pd.read_sql_query(
        """SELECT code AS "Code", pdf_filename AS "PDF File Name", pdf_path AS "PDF Path",
                  modality AS "Modality", pdf_date AS "PDF Date",
                  matched_dicom_count AS "Matched DICOM Count",
                  CASE parse_ok WHEN 1 THEN 'OK' ELSE 'PARSE FAILED' END AS "Filename Parse",
                  notes AS "Notes"
           FROM pdf_results ORDER BY code, pdf_filename""",
        conn,
    )

    folders_df = pd.read_sql_query(
        """SELECT code AS "Code", status AS "Status", error AS "Error",
                  dicom_total AS "Total DICOM Found",
                  dicom_unmatched_no_modality AS "Unmatched (No Modality PDF)",
                  dicom_unmatched_ambiguous AS "Unmatched (Ambiguous Date)",
                  non_dicom_files_skipped AS "Non-DICOM Files Skipped",
                  processed_at AS "Processed At"
           FROM folders ORDER BY code""",
        conn,
    )

    errors_df = folders_df[folders_df["Status"] == "error"]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pdf_df.to_excel(writer, sheet_name="PDF-DICOM Matches", index=False)
        folders_df.to_excel(writer, sheet_name="Folder Summary", index=False)
        errors_df.to_excel(writer, sheet_name="Folder Errors", index=False)

    logger.info("Excel report written to %s", output_path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--folder-a", required=True, help="Path to Folder A (contains PDFs)")
    parser.add_argument("--folder-b", required=True, help="Path to Folder B (contains DICOMs)")
    parser.add_argument("--output", default="pdf_dicom_report.xlsx", help="Output Excel path")
    parser.add_argument(
        "--checkpoint",
        default=None,
        help="SQLite checkpoint DB path (default: <output>.checkpoint.db)",
    )
    parser.add_argument("--workers", type=int, default=4, help="Number of parallel processes")
    parser.add_argument("--force", action="store_true", help="Reprocess codes even if already marked done")
    parser.add_argument("--limit", type=int, default=None, help="Only process the first N codes (for testing)")
    parser.add_argument(
        "--export-only",
        action="store_true",
        help="Skip scanning; just (re)export the Excel report from the existing checkpoint DB",
    )
    parser.add_argument("--log-file", default=None, help="Optional log file path")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.StreamHandler()] + ([logging.FileHandler(args.log_file)] if args.log_file else []),
    )

    output_path = Path(args.output)
    checkpoint_path = Path(args.checkpoint) if args.checkpoint else output_path.with_suffix(".checkpoint.db")
    conn = init_checkpoint_db(checkpoint_path)

    if args.export_only:
        export_excel(conn, output_path)
        return

    dir_a = Path(args.folder_a)
    dir_b = Path(args.folder_b)
    if not dir_a.is_dir():
        parser.error(f"--folder-a is not a directory: {dir_a}")
    if not dir_b.is_dir():
        parser.error(f"--folder-b is not a directory: {dir_b}")

    logger.info("Scanning folder A for exam codes...")
    codes_a = find_code_dirs(dir_a)
    logger.info("Scanning folder B for exam codes...")
    codes_b = find_code_dirs(dir_b)

    common_codes = sorted(set(codes_a) & set(codes_b))
    missing_in_b = sorted(set(codes_a) - set(codes_b))
    missing_in_a = sorted(set(codes_b) - set(codes_a))
    if missing_in_b:
        logger.warning("%d codes present in A but not in B (skipped): %s", len(missing_in_b), missing_in_b[:10])
    if missing_in_a:
        logger.warning("%d codes present in B but not in A (skipped): %s", len(missing_in_a), missing_in_a[:10])

    if args.limit:
        common_codes = common_codes[: args.limit]

    done_codes = set() if args.force else get_done_codes(conn)
    todo_codes = [c for c in common_codes if c not in done_codes]

    logger.info(
        "%d matched codes total, %d already done, %d to process now (workers=%d)",
        len(common_codes), len(done_codes), len(todo_codes), args.workers,
    )

    if not todo_codes:
        logger.info("Nothing to do. Exporting report from checkpoint DB.")
        export_excel(conn, output_path)
        return

    completed = 0
    with concurrent.futures.ProcessPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(process_folder, code, str(codes_a[code]), str(codes_b[code])): code
            for code in todo_codes
        }
        for future in concurrent.futures.as_completed(futures):
            code = futures[future]
            try:
                result = future.result()
            except Exception as exc:  # noqa: BLE001
                result = FolderResult(code=code, status="error", error=f"worker crashed: {exc}")

            save_folder_result(conn, result)
            completed += 1
            if result.status == "done":
                logger.info(
                    "[%d/%d] %s done: %d PDFs, %d DICOMs (%d unmatched-no-modality, %d ambiguous)",
                    completed, len(todo_codes), code, len(result.pdf_records),
                    result.dicom_total, result.dicom_unmatched_no_modality, result.dicom_unmatched_ambiguous,
                )
            else:
                logger.error("[%d/%d] %s FAILED: %s", completed, len(todo_codes), code, result.error)

    export_excel(conn, output_path)


if __name__ == "__main__":
    main()
