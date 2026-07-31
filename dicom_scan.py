import sys
from pathlib import Path

import pydicom
from openpyxl import Workbook


def read_dicom(path):
    try:
        ds = pydicom.dcmread(str(path), stop_before_pixels=True, force=False)
    except Exception:
        return None
    return (
        str(ds.get("StudyDate", "")),
        str(ds.get("Modality", "")),
    )


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    folder = Path(sys.argv[1]).expanduser().resolve()
    if not folder.is_dir():
        print(f"Not a folder: {folder}")
        sys.exit(1)

    output = Path(sys.argv[2]) if len(sys.argv) > 2 else folder / "dicom_report.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "DICOM"
    ws.append(["File Name", "Full Path", "Study Date", "Modality"])

    scanned = 0
    found = 0
    for path in sorted(folder.rglob("*")):
        if not path.is_file():
            continue
        scanned += 1
        info = read_dicom(path)
        if info is None:
            continue
        found += 1
        ws.append([path.name, str(path), info[0], info[1]])
        print(f"[{found}] {path}  {info[0]}  {info[1]}")

    for column, width in zip("ABCD", (40, 90, 14, 12)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"

    wb.save(output)
    print(f"\nScanned {scanned} files, found {found} DICOM files.")
    print(f"Report saved to: {output}")


if __name__ == "__main__":
    main()
