# Master inventory

Answers "what exactly do we have for each patient?" across a large, partly
sorted archive, and writes it to one Excel workbook.

Built for the awkward case: ~5M files and ~1TB on a network share, where the
DICOM headers carry a study date and modality but **no patient identifier**, so
the folder path is the only link between an image and a patient.

## Why three passes

Each pass writes to a SQLite database that the next one reads.

| Pass | Script | Opens files? | Cost | Re-run cost |
|---|---|---|---|---|
| 1 | `crawl.py` | no | hours | resumes |
| 2 | `probe.py` | a sample | ~1 hour | resumes |
| 3 | `report.py` | no | seconds | free |

The split exists because **the code-matching rules will be wrong on the first
try.** Every archive has a naming case nobody anticipated. Keeping attribution
in pass 3 means fixing a rule costs seconds, not another walk of the share.

The file-level inventory lives in SQLite, not Excel, because Excel caps at
~1.05M rows and you have several million files. Only aggregates go to the
workbook; `report.py --csv` dumps the full detail if you need it.

## Two optimisations that matter on a NAS

**One round trip per directory, not per file.** On Windows, `os.scandir`
returns size and timestamps as part of the directory listing, so pass 1 never
stats a file individually. ~200k directory reads instead of ~5M file reads.

**Sample DICOM headers, don't read them all.** Slices in one directory are
almost always one study, so `probe.py` reads a few per directory (always
including the first and last by name) and infers the rest. When the samples
disagree it reads that directory in full and says so. On the test archive this
cut 2,746 file opens to 273. Every study row is marked `inferred` or `counted`
so you can always tell which you're looking at.

## Running it

```bash
pip install -r ../requirements.txt

# 0. Find out what you're in for before committing to the full walk.
python crawl.py "\\nas\studies" --db inventory.db --workers 32 --probe 2000

# 1. Walk everything.
python crawl.py "\\nas\studies" --db inventory.db --workers 32 --expect-files 5000000

# 2. Identify files and extract DICOM metadata.
python probe.py --db inventory.db --workers 32

# 3. See which code prefixes actually exist, THEN attribute and report.
python report.py --db inventory.db --discover
python report.py --db inventory.db --prefixes AA,AVDD,QQQ --out master_report.xlsx

# 4. Optional: the same facts in the old match_reports.py workbook format.
python match_report_from_db.py --db inventory.db --out match_report.xlsx
```

All three resume if interrupted - re-run the same command. Raise `--workers` on
a network share (latency-bound, so parallelism helps a lot); lower it on a
local disk.

Long Windows paths (>260 chars) are handled via the `\\?\` prefix. Without it
an over-long directory silently reads as empty, which is the worst kind of bug
to have in an inventory.

## The workbook

| Sheet | One row per | Answers |
|---|---|---|
| `Overview` | - | headline counts, file-type totals |
| `Patients` | patient code | file counts by type, size, studies with/without a report |
| `Studies` | study | date, modality, slice count, matched report, match tier |
| `Unassigned` | folder | files with no derivable code - **the work queue** |
| `Coverage` | patient code | images / reports / both, and what is missing |
| `Breakdown` | patient code x category | the slices, modalities and dates behind each |
| `Suggestions` | orphan study | candidate codes recovered by cross-referencing reports |
| `Conflicts` | file | filename and folder disagree about the code |

`Unassigned` and `Conflicts` are the point. The counts tell you what you have;
those two tell you what is still broken.

### Suggestions

A study in an unnamed folder cannot be identified from its files - there is no
PatientID in the header. But reports are named `<code> <modality> <date>.pdf`,
and DICOM headers carry date and modality, so that overlap can point back to a
code. If exactly one patient has a report on that date and modality, that is a
strong lead. It is a lead, not a match, which is why it sits on its own sheet.

### Status and Note

`Status` on the `Studies` sheet is one of three plain verdicts, so it can be
filtered on without knowing every phrasing of the reason:

| Status | |
|---|---|
| `report and image` | both present |
| `only image` | images, no report paired with them |
| `only report` | a report whose images are missing |
| `other pdf` | a PDF with no date in its name - a consent form, a manual, `HELP.pdf`. Listed, but not chased |

(`no report or image` completes the set, but no *study* row can be it - a
patient with neither appears on `Coverage`.)

**Why** sits in `Note`, next to it:

```
 13  report and image   | matched on date + modality
  1  report and image   | matched on date only - modality not confirmed
  5  only report        | no images found for this report
  4  only image         | folder carries no code - see Suggestions
  3  only image         | patient has other reports, none for this date
  2  other pdf          | no date in the file name - probably not a study report
  1  only image         | no report anywhere for this patient
```

Not every PDF in a patient folder is a report. A report carries a date in its
name; a consent form or a leaflet does not, and so can never match a study.
Calling those `only report / no images found` sends someone hunting for images
that were never meant to exist, so they get their own status and are kept out
of the report counts - `Coverage` shows them in an `Other PDFs` column, and a
patient whose only PDF is one of these is not counted as reported.

A large `matched on date only` count means the modality word in those report
filenames could not be mapped - check `MODALITY_ALIASES` in
`../match_reports.py`. Testing surfaced exactly this: `XR` is absent, so every
X-ray silently drops to a date-only match. `report.py` imports that table
rather than duplicating it, so a fix there applies to both tools.

## Who has what: the `Coverage` sheet

A patient is rarely all-or-nothing, so one label per patient is not enough.
`AA0006` below has a study with a report, a second study with none, **and** a
report whose images never arrived - all three at once. Every row carries the
counts behind its label:

```
Patient Code  DICOM  Report Files  Other PDFs  Studies  Report and Image  Only Image  Only Report  Category
AA0050            2             0           0        1                 0           1            0  only image
AA0051            0             1           0        0                 0           0            1  only report
AA0006           52             2           1        2                 1           1            1  report and image (partial)
AA0009           12             1           0        1                 1           0            0  report and image (complete)
AA0052            0             0           0        0                 0           0            0  no report or image
```

Every count column is named for the `Status` value it counts. `Only Image` = 1
for AA0006 means filtering `Status = only image` on the `Studies` sheet finds
exactly that one row - the number and the rows behind it share a phrase, so
there is nothing to translate between sheets.

`Studies` = `Report and Image` + `Only Image`. The other two columns count
**PDFs**, not studies, which is why they sit outside that total: a row with no
images cannot be a study. On the `Studies` sheet those rows show `DICOM
Files = 0`.

| label | means |
|---|---|
| `only image` | no report anywhere for this patient |
| `only report` | reported, but no DICOM file is attributed to them |
| `report and image (partial)` | has both, but some study has no report or some report has no images - **the three columns say which** |
| `report and image (complete)` | every study has a report and every report has images |
| `no report or image` | a code folder holding neither - only Word/Excel files, stray litter, or nothing at all |

Sort or filter on `Category` for the roll-up; read the columns for the detail.
The counts also appear on `Overview`, and `match_report_from_db.py` prints them
at the end of every run, with the study-level totals underneath:

```
  patients: report and image (partial)        6   42.9%
  patients: report and image (complete)       5   35.7%
  ...
  rows: report and image                     14
  rows: only image                            4
  rows: only report                           5
  rows: other pdf                             2
```

Add `--coverage patients.xlsx` there for the per-patient detail as its own
workbook. It is a separate file on purpose: `match_report.xlsx` reproduces
`match_reports.py` exactly, and anything reading it expects one sheet with six
known columns.

The per-study verdicts come from each tool's own matching rather than from
`coverage.py`, so a coverage row can never contradict the sheet next to it -
`AA0006` above is the same three rows the workbook shows for that patient.

**"Has images" means at least one file identified as DICOM, not at least one
study.** A folder whose headers could not be parsed still holds images;
counting studies would quietly file it under "no images".

## The `Breakdown` sheet: what is actually on each side

`Coverage` says AA0006 is partly covered. `Breakdown` says which studies are on
which side of that, without opening `Studies` and filtering by hand:

```
Patient Code  Category           Studies  DICOM Files  Modalities  Dates
AA0003        report and image         1           90  Echo        20170219
AA0003        only image               1           40  Echo        20151007
AA0006        report and image         1           12  MRI         20170206
AA0006        only image               1           40  X-Ray       20220206
AA0006        only report              1            0              20230523
AA0006        other pdf                1            0
AA0012        report and image         2           52  CT, MRI     20180427, 20190912
```

One row per patient per category, so a patient appears once for each side they
have something on. `DICOM Files` is 0 on the two PDF categories by definition -
there are no images behind them.

Patients with a long history get their dates capped at 12 followed by
`(+N more)`, which keeps the cell readable and well inside Excel's limit.

`report.py` writes it as a sheet; `match_report_from_db.py --coverage` writes
the same sheet into its coverage workbook. Both are built by `coverage.py` from
each tool's own matching, and the two come out identical.

## `patient_summary.py`: one line per patient

A second, smaller workbook for the question "what exactly does AA0001 have?" -
just two sheets, no study-level detail:

```bash
python patient_summary.py --db inventory.db --out patient_summary.xlsx
```

```
Patient  Total DICOM  Total  DICOMs With  DICOMs With  Reports With  Other  Total Files  Total
Code          Images   PDFs    a Report     No Report      No DICOM   PDFs  (All Types)   Size
AA0003           130      1          90            40             0      0          131  27.7 KB
AA0006            52      3          12            40             1      1           58  11.2 KB
AVDD004           92      2          52            40             0      0           96  19.7 KB
```

Headings are spelled out rather than abbreviated, because this is the file that
gets forwarded to someone who does not work with the archive daily.

**These are DICOM *files*, not studies.** `master_report.xlsx` counts studies,
because that is the unit a report is written about; this counts the files
underneath them, because that is the unit the archive is measured in. A patient
with one reported study of 400 slices and one unreported study of 200 appears
here as 400 and 200, and there as 1 and 1. Both are right - this is the one
people mean when they ask how much is outstanding.

`Reports Covering Them` is how many distinct PDFs account for the reported
side: one report can cover 400 slices, or several can.

Totals come first, then the split, so the parts can be checked against the
whole at a glance:

* `DICOMs With a Report` + `DICOMs With No Report` = `Total DICOM Images`
* `Total Files (All Types)` counts everything attributed to the patient -
  images, PDFs, Word, video, anything else - so the difference from the two
  totals beside it is the non-image non-PDF material.

If the first of those does not add up, the missing files are DICOMs that no
study accounts for, usually a header `probe.py` could not read. `Overview`
carries that number as `Image files not in any study`; it is normally 0.

Rows with anything outstanding - DICOMs with no report, or reports with no
DICOM - are shaded.

It reads the matching `report.py` already wrote to `study_report`, so nothing
is re-matched and the two files cannot disagree - which does mean `report.py`
has to have been run first.

## The old `match_reports.py` format

Whatever already reads `match_report.xlsx` and `results/<code>-results.xlsx`
keeps working - `match_report_from_db.py` rebuilds exactly those files from the
database instead of walking the share again:

```bash
python match_report_from_db.py --db inventory.db --out match_report.xlsx
```

Sheet name, header, column widths, frozen header row, status strings and
per-code file naming are imported from `match_reports.py`, not copied, so the
two cannot drift. It runs in seconds, so a fix to the matching rules is a
re-run, not another crawl.

Three things differ from running `match_reports.py` itself, all in your favour
except the last:

* **Reports are found anywhere in the archive**, not just under a matching
  folder in a separate reports root. On the test archive this turned nine
  "no report found" rows into matches - the PDFs were sitting in the image
  folders.
* **Codes are normalised**, so `AA006` and `AA0006` are one patient with one
  file, not two.
* **DICOM File Count comes from the `studies` table.** With `probe.py
  --metadata all` (the default) it is exact; with `sample` or `none` some
  counts are inferred, and the script prints how many rows that affects.

One prefix at a time, without disturbing the attribution in the database:

```bash
python match_report_from_db.py --db inventory.db --prefix AA \
                               --out AA_match_report.xlsx --results-dir AA_results
```

Filter here, not by re-running `report.py --prefixes AA` - that rewrites
attribution for the whole archive and would drop every AVDD and QQQ code on the
floor. Attribute all the real prefixes once; narrow at this step. `--code
AA0006,AA0012` narrows to exact codes the same way.

Studies in folders with no code, and PDFs with no derivable code, have no
per-code workbook to live in; they are counted in the run summary and belong on
the master workbook's `Unassigned` sheet. `--include-unassigned` folds the
studies into the combined summary anyway, with an empty `Code`.

## Prefixes: always run `--discover` first

Prefixes vary by study (`AA`, `AVDD`, `QQQ`, and whatever the next one is), so
there is no safe default pattern. A generic `[A-Z]{2,6}\d{1,4}` looks
reasonable and is a trap - on the test archive it matches `IM00011` (DICOM
slice names), `batch 1`, `study0` and `reviewed 2024`, inventing thousands of
patients that don't exist.

So `--discover` lists every code-like token with counts and an example, you
read it, and you pass the real ones to `--prefixes`.

## When the codes come out wrong

`diagnose_codes.py` reads the database and says why. It writes nothing:

```bash
python diagnose_codes.py --db inventory.db --prefixes AA
```

**A name like `EEAA6079` or `AA1234AA` invented a patient.** `--prefixes AA`
means the code *is* `AA` plus digits, not a fragment of a longer token, so the
pattern is guarded at both ends: a prefix that starts the token, a digit run,
then something that is neither a letter nor a digit.

```
EEAA6079             no match      AA1234AA            no match
XAA0001              no match      AA 20240115 rescan  no match
2024AA0001           no match      AA123456789         no match

AA0001               AA0001        _AA0001             AA0001
AA0001 follow up     AA0001        scan-AA0001         AA0001
AA0001.pdf           AA0001        AVDD1400            AVDD1400
```

Separators are not letters or digits, so `_AA0001` and `scan-AA0001` still
resolve - only a name running letters or digits straight into the code is
rejected. A file named `EEAA6079` inside `AA6554 base` now takes `AA6554` from
its folder with `code_source = folder`, instead of inventing `AA6079` and
raising a filename/folder conflict against its own patient.

If your codes are always a fixed width, `--digits 4` says so outright and
refuses anything else; `--digits 3-5` takes a range. The default is 1-5.

**Every later tool displays these codes rather than deriving its own**, so
after changing a matching rule, `report.py` has to be re-run before
`patient_summary.py` or `match_report_from_db.py` will show the difference -
until then they faithfully print the previous run's attribution. Both now open
with the stamp of the run that wrote what they are showing:

```
codes written by report.py at 2026-09-16 14:27:02  (--prefixes AA --digits 4)
```

**Codes padded too wide** (`AA0001` coming out as `AA000001`) was a bug, fixed
in two places. A folder called `AA 20240115 rescan` used to match as `AA20240`:
an invented patient whose long number then repadded every real code. Now the
pattern ends in `(?!\d)`, so a code is the whole digit run or nothing, and that
name is left unassigned where it belongs. Padding is also learned from the
*most common* width rather than the widest, so a few odd names can no longer
repad thousands of real ones. Longer numbers are never truncated - the width is
a minimum, so a genuine `AA12345` still prints in full.

`--pad 4` (or `--pad AA=4,AVDD=3`) forces the width if you still need to.

**`--discover` doesn't list a prefix that is obviously there.** The two patterns
are not the same: `--discover` is word-bounded on both sides, while
`--prefixes AA` builds `(?:AA)[-_ ]?\d{1,5}(?![A-Za-z0-9])`. So
`SCANAA0001` is invisible to `--discover` and still matched by `--prefixes` -
which is why a prefix can be missing from the table and attribute thousands of
files anyway. Section 4 of the diagnosis shows which names fall in that gap.

**The Excel file shows codes the database doesn't have.** `report.py` clears
every code at the start of each run, so only the last run survives in the
database. A workbook from an earlier run, or one built against a different
`--db`, will disagree with it. Section 1 prints the database's root and row
counts so you can tell which one you are looking at.

## Normalising codes

`AA001`, `AA_1` and `AA0001` are one patient. Left alone they become three and
the error is invisible in a total - it just looks like more patients with fewer
files each. Codes are keyed on **prefix + integer**, so padding cannot split
them, and padding width is learned per prefix (`AVDD001` is 3 digits, `AA0001`
is 4; a single global setting would corrupt one of them). The width is the most
common one seen for that prefix, so a stray name cannot widen the rest, and it
is a minimum rather than a fixed size - a longer number still prints in full.

## Querying the database directly

Attribution is written back to the database, so you don't have to go through
Excel. The `v_files` view joins everything together:

```sql
-- everything belonging to one patient, whatever the folder was called
SELECT kind, COUNT(*) FROM v_files WHERE code = 'AA0006' GROUP BY kind;

-- every DICOM file for a patient, full paths
SELECT full_path FROM v_files WHERE code = 'AA0006' AND kind = 'dicom';

-- the work queue: files nothing could link to a patient
SELECT kind, COUNT(*) FROM v_files WHERE code IS NULL GROUP BY kind;

-- where did the code come from - the file name, or the folder?
SELECT code_source, COUNT(*) FROM v_files GROUP BY code_source;
```

## Full DICOM metadata

`probe.py --metadata` decides how much header data is kept. The default is
`all`: read every DICOM file and store its complete header as DICOM JSON in
`dicom_meta`, so that any field anyone asks for later is a local query rather
than another crawl of the share.

| mode | files opened | DB size (~3.5M DICOM) | study counts |
|---|---|---|---|
| `all` (default) | every DICOM | ~20 GB | exact |
| `sample` | a few per directory | ~300 MB | inferred |
| `none` | a few per directory | tiny | inferred |

`all` is the default, and the only mode that gives exact study counts - the
other two infer most of them from a sample. Every study row is marked
`counted` or `inferred`, `Overview` totals both, and `report.py` prints a
warning if any study count is inferred, so an estimate can never be mistaken
for a real number.

**A resumed probe skips directories already done, whatever `--metadata` now
says.** Switching from `sample` to `all` on a probed database therefore
changes nothing on its own - it reports `0 directories to probe` and keeps
every inferred count. `probe.py` warns when the mode differs from the one the
database was built with; to actually re-read, add `--force`:

```bash
python probe.py --db inventory.db --metadata all --force
```

Check what you have at any time:

```sql
SELECT confidence, COUNT(*) FROM studies GROUP BY confidence;
```

`all` is slower and much bigger, and worth it if the archive gets asked
questions more than once. Time the share first:

```bash
python probe.py --db inventory.db --workers 32 --probe 500
```

That stops after 500 directories and reports files/sec and bytes/file, so you
can extrapolate before committing. It resumes where it stopped.

JSON is stored uncompressed on purpose - compressing it would halve the size
but break `json_extract`, which defeats the point of keeping it.

Query it with SQLite's JSON functions:

```sql
SELECT json_extract(json, '$."00080060".Value[0]') AS modality,
       json_extract(json, '$."00080020".Value[0]') AS study_date,
       COUNT(*)
  FROM dicom_meta GROUP BY 1, 2;
```

Or use `extract.py`, which does the joining and tag-name lookup for you:

```bash
# what tags do we actually have?
python extract.py --db inventory.db --list-tags

# one row per study, with patient code and folder
python extract.py --db inventory.db --tags StudyDate,Modality --by study \
                  --out fields.csv

# keywords and raw hex tags mix freely; --by file for per-slice detail
python extract.py --db inventory.db --tags Manufacturer,00181030 --by file \
                  --out fields.csv
```

**This table contains PHI** - patient name, DOB, institution - so the database
belongs wherever the archive itself is allowed to live.

## Testing

`make_fake_archive.py` builds a deliberately messy archive - extensionless
DICOMs, inconsistent padding, studies filed under scan-date folders, reports
split between patient folders and a separate tree, plus the usual litter. It
writes real parseable DICOM headers, so the whole pipeline can be exercised
without touching patient data.

```bash
python make_fake_archive.py /tmp/fake --patients 40
python crawl.py /tmp/fake --db /tmp/inv.db && python probe.py --db /tmp/inv.db
python report.py --db /tmp/inv.db --out /tmp/master.xlsx
```
