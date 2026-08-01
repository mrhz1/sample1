import sys
from pathlib import Path

from pypdf import PdfReader
from openpyxl import Workbook


def read_pdf(path):
    try:
        reader = PdfReader(str(path))
        info = reader.metadata or {}
    except Exception:
        return None
    return (
        str(info.get("/Title", "")),
        str(info.get("/Author", "")),
        str(info.get("/Subject", "")),
        str(info.get("/Creator", "")),
        str(info.get("/Producer", "")),
        str(info.get("/CreationDate", "")),
        str(info.get("/ModDate", "")),
        len(reader.pages),
    )


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    folder = Path(sys.argv[1]).expanduser().resolve()
    if not folder.is_dir():
        print(f"Not a folder: {folder}")
        sys.exit(1)

    output = Path(sys.argv[2]) if len(sys.argv) > 2 else folder / "pdf_report.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "PDF"
    ws.append([
        "File Name", "Full Path", "Title", "Author", "Subject",
        "Creator", "Producer", "Creation Date", "Mod Date", "Pages",
    ])

    scanned = 0
    found = 0
    for path in sorted(folder.rglob("*.pdf")):
        if not path.is_file():
            continue
        scanned += 1
        info = read_pdf(path)
        if info is None:
            continue
        found += 1
        ws.append([path.name, str(path), *info])
        print(f"[{found}] {path}  {info[0]}  pages={info[7]}")

    for column, width in zip("ABCDEFGHIJ", (30, 60, 25, 20, 20, 20, 20, 22, 22, 8)):
        ws.column_dimensions[column].width = width
    ws.freeze_panes = "A2"

    wb.save(output)
    print(f"\nScanned {scanned} PDF files, found {found} readable PDF files.")
    print(f"Report saved to: {output}")


if __name__ == "__main__":
    main()
