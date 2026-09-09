"""Build a small, deliberately messy archive for testing the inventory scripts.

Nothing here touches real data - it exists so the crawler and the analysis can
be exercised end to end before they ever run on the NAS. The messiness is
modelled on the described archive:

  - sorted areas where a patient folder holds everything for that patient
  - unsorted dumps where studies sit under scan-date or operator folders and
    the code appears only in a filename, or nowhere at all
  - DICOM files with no extension at all
  - inconsistent zero padding (AA001 vs AA0001)
  - reports named "<code> <modality> <date>.pdf", some inside the patient
    folder, some in a separate reports tree
  - the usual archive litter: viewer executables, thumbs.db, readmes

Usage:
    python make_fake_archive.py <target_dir> [--patients N] [--seed N]
"""

import argparse
import os
import random
import shutil
import struct

# (word used in report filenames, code stored in the DICOM header)
MODALITIES = [("CT", "CT"), ("MRI", "MR"), ("Echo", "US"),
              ("XR", "CR"), ("US", "US")]
LITTER = ["Thumbs.db", "readme.txt", "DICOMDIR", "VIEWER.EXE", "autorun.inf"]


def write_dicom(path, study_date, modality, study_uid):
    """A minimal but genuinely parseable DICOM: 128-byte preamble, 'DICM',
    then explicit-VR little-endian elements for the three tags the analysis
    actually reads."""
    out = bytearray(b"\0" * 128 + b"DICM")

    def elem(group, elem_no, vr, value):
        data = value.encode("ascii")
        if len(data) % 2:
            data += b" "
        return struct.pack("<HH2sH", group, elem_no, vr, len(data)) + data

    # (0002,0010) TransferSyntaxUID, preceded by the group-length element the
    # meta group requires: VR "UL", length 4, value = bytes that follow.
    ts = elem(0x0002, 0x0010, b"UI", "1.2.840.10008.1.2.1")
    out += struct.pack("<HH2sH", 0x0002, 0x0000, b"UL", 4)
    out += struct.pack("<I", len(ts))
    out += ts
    # Dataset, explicit VR little endian per the transfer syntax above.
    out += elem(0x0008, 0x0020, b"DA", study_date)
    out += elem(0x0008, 0x0060, b"CS", modality)
    out += elem(0x0020, 0x000D, b"UI", study_uid)
    with open(path, "wb") as fh:
        fh.write(out)


def write_pdf(path, text):
    with open(path, "wb") as fh:
        fh.write(b"%PDF-1.4\n% fake report: " + text.encode() + b"\n%%EOF\n")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("target")
    ap.add_argument("--patients", type=int, default=40)
    ap.add_argument("--seed", type=int, default=7)
    args = ap.parse_args()

    rnd = random.Random(args.seed)
    root = os.path.abspath(args.target)
    if os.path.exists(root):
        shutil.rmtree(root)
    os.makedirs(root)

    sorted_root = os.path.join(root, "Sorted - reviewed 2024")
    dump_root = os.path.join(root, "TO SORT")
    reports_root = os.path.join(root, "Reports backup")
    for d in (sorted_root, dump_root, reports_root):
        os.makedirs(d)

    stats = {"dicom": 0, "pdf": 0, "other": 0}

    # Several studies share the archive, each with its own prefix and its own
    # idea of how many digits a code has. AA is additionally inconsistent with
    # itself, which is the case that silently splits one patient into two.
    STUDIES = [("AA", 4), ("AVDD", 3), ("QQQ", 3)]

    for n in range(1, args.patients + 1):
        prefix, width = STUDIES[n % len(STUDIES)]
        canonical = f"{prefix}{n:0{width}d}"
        # The same patient spelled two ways in two places - images under
        # AA003, reports under AA0003. This is the case that silently splits
        # one patient into two if codes aren't normalised on the number.
        code = canonical
        alt_code = canonical
        if prefix == "AA" and n % 2 == 0:
            alt_code = f"{prefix}{n:03d}"
        n_studies = rnd.choice([1, 1, 2, 3])

        for s in range(n_studies):
            mod_word, mod_code = rnd.choice(MODALITIES)
            date = f"20{rnd.randint(15, 23)}{rnd.randint(1, 12):02d}{rnd.randint(1, 28):02d}"
            pretty = f"{date[6:8]}-{date[4:6]}-{date[0:4]}"
            uid = f"1.2.826.{n}.{s}"
            n_slices = rnd.choice([12, 40, 90])

            placement = rnd.random()
            if placement < 0.55:
                # Sorted: everything under one patient folder.
                suffix = rnd.choice(["", "", " follow up", " base date"])
                study_dir = os.path.join(
                    sorted_root, code + suffix, f"{mod_word} {pretty}", "DICOM")
                report_dir = os.path.join(sorted_root, code + suffix)
            elif placement < 0.8:
                # Unsorted: code is in the path but buried under a dump folder.
                study_dir = os.path.join(
                    dump_root, f"batch {rnd.randint(1, 4)}",
                    f"{code} {mod_word.lower()}", "IMG")
                report_dir = os.path.join(reports_root, alt_code)
            else:
                # Worst case: no code anywhere in the path. Only the report,
                # filed elsewhere, knows who this is.
                study_dir = os.path.join(
                    dump_root, "unsorted scans", f"scan {date}", f"study{s}")
                report_dir = os.path.join(reports_root, alt_code)

            os.makedirs(study_dir, exist_ok=True)
            for i in range(n_slices):
                # No extension - the common real-world case.
                write_dicom(os.path.join(study_dir, f"IM{i:05d}"), date, mod_code, uid)
                stats["dicom"] += 1

            # Most studies have a report; some deliberately don't.
            if rnd.random() < 0.75:
                os.makedirs(report_dir, exist_ok=True)
                sep = rnd.choice([" ", "-"])
                write_pdf(
                    os.path.join(
                        report_dir,
                        f"{alt_code}{sep}{pretty}{sep}{mod_word}.pdf"),
                    f"{canonical} {mod_word}")
                stats["pdf"] += 1

        # Litter and the odd loose office file.
        patient_dir = os.path.join(sorted_root, code)
        if os.path.isdir(patient_dir):
            for name in rnd.sample(LITTER, rnd.randint(0, 3)):
                open(os.path.join(patient_dir, name), "w").write("x")
                stats["other"] += 1
            if rnd.random() < 0.4:
                open(os.path.join(patient_dir, f"{code} notes.docx"), "w").write("x")
                stats["other"] += 1

    # A couple of orphans with no patient link at all.
    misc = os.path.join(root, "misc")
    os.makedirs(misc)
    for name in ("old backup.zip", "scanner log.txt", "install.exe"):
        open(os.path.join(misc, name), "w").write("x")
        stats["other"] += 1

    print(f"built {root}")
    print(f"  {args.patients} patients, {stats['dicom']:,} DICOM, "
          f"{stats['pdf']:,} PDF, {stats['other']:,} other")


if __name__ == "__main__":
    main()
