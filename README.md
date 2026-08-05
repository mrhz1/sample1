# DICOM / PDF report matching toolkit

A small set of standalone Python scripts for reconciling an archive of DICOM
studies against a folder of PDF radiology reports, plus a few utilities for
inspecting the raw files along the way.


The main script is **`match_reports.py`** — it answers "which PDF report
belongs to which set of images?" across an archive of millions of files, and is
built to survive being interrupted partway through a multi-day run.

## Setup

Python 3.10 (see `.python-version`).

```bash
python -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

The codec packages (`pylibjpeg*`, `python-gdcm`) are only needed by
`dicom_view.py` for rendering compressed pixel data; matching and scanning read
headers only and don't touch them.

## Expected folder layout

`match_reports.py` assumes two trees whose subfolders share the same R-code
naming:

```
images/
  R123/...            DICOM files, at any nesting depth, mixed in with
  R456/...            other junk (viewer apps, readmes, thumbnails)
reports/
  R123/R123 CT 10-10-2010.pdf
  R456/R456 MRI 01-15-2014.pdf
```

PDF filenames are expected to contain a date as `MM-DD-YYYY` (`-`, `_`, or `.`
separators) and, ideally, a modality word. Files that don't match the pattern
aren't dropped — they just fall to a weaker matching tier.

---

## `match_reports.py` — match reports to studies

```bash
python match_reports.py <images_folder> <reports_folder> [output.xlsx] \
                        [--rcode CODE[,CODE...]] [--force] [--debug]
```

### How matching works

Tiered, strongest first:

1. **R-code** — the shared subfolder name. Required; nothing matches across
   different R-codes.
2. **Study date + modality** — the DICOM `StudyDate` and `Modality` tags vs. the
   date and modality parsed out of the PDF filename. This tier is what keeps a
   CT and an Echo done on the *same day* from being conflated.
3. **Study date only** — fallback when the modality can't be parsed from the
   filename or is missing from the DICOM tag. Flagged in the output as
   "modality unconfirmed".

Anything weaker is left as "review manually" rather than guessed.

Two lookup tables near the top of the file are meant to be extended as you meet
real-world data: `MODALITY_ALIASES` (filename word → DICOM code, e.g.
`ECHO`→`US`) and `MODALITY_DISPLAY_NAMES` (DICOM code → label shown in Excel,
e.g. `MR`→`MRI`). Display names are cosmetic only; matching always uses raw
DICOM codes.

### Output

```
match_report.xlsx           every R-code combined, rebuilt at the end of each run
results/
  R123-results.xlsx         one file per R-code, written the moment it finishes
  R456-results.xlsx
  progress.log              append-only timestamped log of the whole run
```

Every sheet has the same columns: `Code`, `Study Date`, `Modality`,
`DICOM File Count`, `Matched PDF File(s)`, `Match Status`. One row per study
(R-code + date + modality), followed by rows for PDFs that matched no images
(`No DICOM images found`).

### Resuming an interrupted run

Any R-code that already has a `results/<code>-results.xlsx` is skipped, so
after a crash, a dropped connection, or a Ctrl-C you re-run **the exact same
command** and it continues where it stopped:

```bash
python match_reports.py /mnt/archive/images /mnt/archive/reports
# ... dies at R2891 ...
python match_reports.py /mnt/archive/images /mnt/archive/reports   # picks up at R2891
```

Per-code files are written to a temp file and atomically renamed, so a process
killed mid-write leaves no half-finished file that would be mistaken for done.
`match_report.xlsx` is rebuilt from everything in `results/` at the end of every
run, so it stays complete across resumes — including after a `--rcode` run.

Use `--force` to re-scan R-codes that already have a results file.

### Watching progress

All progress goes to stdout *and* to `results/progress.log`, flushed per line:

```bash
tail -f results/progress.log
```

```
[2026-08-05 09:14:02] 450 R-code folder(s) in scope; 0 already have results, 450 to do
[2026-08-05 09:14:02] [1/450] R123: starting (449 folder(s) left after this)
[2026-08-05 09:18:12]     R123: 20,000 files examined, 18,412 DICOM, 4m10s elapsed (80 files/s)
[2026-08-05 09:20:04] [1/450] R123: done in 6m02s - 28,411 files (26,003 DICOM), 2 PDF(s),
                      3 matched study bucket(s), 0 needing review, 0 unmatched PDF(s)
[2026-08-05 09:20:04]     run elapsed 6m02s, rough ETA 45h11m for the rest
```

The in-folder heartbeat fires every `PROGRESS_EVERY = 2000` files — raise that
constant if it's too chatty at your file counts.

### Running over specific codes only

```bash
python match_reports.py images reports --rcode R123
python match_reports.py images reports --rcode R123,R456,R789
```

Non-listed subfolders are skipped entirely rather than scanned and filtered
afterward, so this is genuinely cheap on a large archive.

### Diagnosing a "should have matched" case

```bash
python match_reports.py images reports --rcode R123 --debug
```

`--debug` prints, per file: `dcmread` failures with the exception message and
whether `force=True` would have worked (the usual cause is a missing DICM
preamble), the date/modality parsed from each PDF filename, and a side-by-side
of R-code folder names using `repr()` so hidden whitespace or case differences
become visible.

### Performance notes

DICOM files are read with `stop_before_pixels=True`, so only headers are
parsed — the scan is I/O-bound on directory traversal, not on image decoding.
Memory stays flat regardless of archive size: state is kept per R-code and
released once that code's file is written. Output rows are per *study*, not per
file, so the ~1,048,576-row Excel limit is not a concern even for millions of
images.

---

## Supporting scripts

### `dicom_scan.py` — inventory every DICOM file in a tree

```bash
python dicom_scan.py <folder> [output.xlsx]
```

Walks a folder recursively and writes one row per readable DICOM file (name,
full path, `StudyDate`, `Modality`). Non-DICOM files are silently skipped.
Useful for seeing what's actually in an archive before matching. Defaults to
`<folder>/dicom_report.xlsx`.

### `pdf_metadata.py` — inventory PDF metadata

```bash
python pdf_metadata.py <folder> [output.xlsx]
```

One row per PDF: title, author, subject, creator, producer, creation/mod dates,
page count. Defaults to `<folder>/pdf_report.xlsx`.

### `dicom_view.py` — look at a DICOM file

```bash
python dicom_view.py <dicom_file> [output.png]
python dicom_view.py <folder> [output_folder]     # converts everything it finds
```

Renders a DICOM to PNG and prints its identifying tags — for spot-checking that
a study is what the tags claim (e.g. that a `US` file really is an
echocardiogram) without a dedicated viewer. Handles multi-frame and color
Doppler exports by collapsing extra axes to a single displayable frame, and
applies a VOI LUT plus percentile windowing to grayscale images.

Files with no pixel data aren't a failure case: Structured Reports get their
content tree printed as readable text, and anything else gets a full raw tag
dump with binary fields summarized rather than spewed as bytes.

If pixel decoding fails on a compressed transfer syntax, that's a missing codec
— `pylibjpeg-openjpeg` and `python-gdcm` in `requirements.txt` cover most cases.

## Known limitations

- Summary rows bucket on `(R-code, study date, modality)`, so two genuinely
  distinct same-day studies of the *same* modality collapse into one row.
- Matching trusts the folder naming convention completely — an R-code folder
  misnamed on one side (trailing space, different case) matches nothing.
  `--debug` exists specifically to surface that.
