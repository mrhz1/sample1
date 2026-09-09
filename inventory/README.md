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

### Match tiers

Reported separately on purpose, so a weak match never looks like a strong one:

1. `matched (date + modality)` - strongest
2. `matched (date only - modality unconfirmed)` - the modality word in the
   report filename could not be mapped
3. `report exists for patient but no date match`
4. `no report found`

If tier 2 is large, check `MODALITY_ALIASES` in `../match_reports.py` for a
missing word. Testing surfaced exactly this: `XR` is absent, so every X-ray
silently drops a tier. `report.py` imports that table rather than duplicating
it, so a fix there applies to both tools.

## Prefixes: always run `--discover` first

Prefixes vary by study (`AA`, `AVDD`, `QQQ`, and whatever the next one is), so
there is no safe default pattern. A generic `[A-Z]{2,6}\d{1,4}` looks
reasonable and is a trap - on the test archive it matches `IM00011` (DICOM
slice names), `batch 1`, `study0` and `reviewed 2024`, inventing thousands of
patients that don't exist.

So `--discover` lists every code-like token with counts and an example, you
read it, and you pass the real ones to `--prefixes`.

## Normalising codes

`AA001`, `AA_1` and `AA0001` are one patient. Left alone they become three and
the error is invisible in a total - it just looks like more patients with fewer
files each. Codes are keyed on **prefix + integer**, so padding cannot split
them, and padding width is learned per prefix (`AVDD001` is 3 digits, `AA0001`
is 4; a single global setting would corrupt one of them).

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
