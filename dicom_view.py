"""
Convert a DICOM file to a viewable PNG, and print its identifying tags.

Useful for spot-checking what a study actually is (e.g. confirming a "US"
modality file really is an echocardiogram and not something else) without
a dedicated DICOM viewer - VS Code can't open .dcm files directly, but it
can preview the PNG this produces.

Usage:
    python dicom_view.py <dicom_file> [output.png]
    python dicom_view.py <folder> [output_folder]   # converts every DICOM file found
"""

import sys
from pathlib import Path

import numpy as np
import pydicom
from PIL import Image

# VRs that hold binary payloads (pixel-like data, compressed streams, unknown
# private data) rather than human-readable text. These must stay summarized
# in the raw dump - printing them in full dumps raw bytes as unreadable junk
# (e.g. a small embedded icon thumbnail) right alongside genuine report text.
BINARY_VRS = {"OB", "OD", "OF", "OL", "OV", "OW", "UN", "OB or OW", "US or SS", "US or SS or OW"}


TAGS_TO_PRINT = [
    "Modality", "StudyDescription", "SeriesDescription", "BodyPartExamined",
    "StudyDate", "SeriesNumber", "InstanceNumber", "Rows", "Columns",
    "NumberOfFrames", "SamplesPerPixel", "PhotometricInterpretation",
    "PlanarConfiguration",
]


def format_sr_item(item):
    """Return (label, value_type, value) for one SR content item."""
    label = ""
    if item.get("ConceptNameCodeSequence"):
        label = item.ConceptNameCodeSequence[0].get("CodeMeaning", "")
    value_type = item.get("ValueType", "")
    value = ""
    if value_type == "TEXT":
        value = item.get("TextValue", "")
    elif value_type == "CODE" and item.get("ConceptCodeSequence"):
        value = item.ConceptCodeSequence[0].get("CodeMeaning", "")
    elif value_type == "NUM" and item.get("MeasuredValueSequence"):
        mv = item.MeasuredValueSequence[0]
        num = mv.get("NumericValue", "")
        units = ""
        if mv.get("MeasurementUnitsCodeSequence"):
            units = mv.MeasurementUnitsCodeSequence[0].get("CodeValue", "")
        value = f"{num} {units}".strip()
    elif value_type == "DATETIME":
        value = item.get("DateTime", "")
    elif value_type == "DATE":
        value = item.get("Date", "")
    elif value_type == "TIME":
        value = item.get("Time", "")
    elif value_type == "UIDREF":
        value = item.get("UID", "")
    elif value_type == "PNAME":
        value = str(item.get("PersonName", ""))
    elif value_type in ("IMAGE", "COMPOSITE", "WAVEFORM") and item.get("ReferencedSOPSequence"):
        ref = item.ReferencedSOPSequence[0]
        value = f"-> references SOPInstanceUID {ref.get('ReferencedSOPInstanceUID', '')}"
    return label, value_type, value


def dump_sr_tree(ds, indent=0):
    for item in ds.get("ContentSequence", []):
        label, value_type, value = format_sr_item(item)
        prefix = "  " * (indent + 1)
        rel = item.get("RelationshipType", "")
        line = f"{prefix}- [{rel}] {label or '(unlabeled)'}" if rel else f"{prefix}- {label or '(unlabeled)'}"
        if value_type:
            line += f" ({value_type})"
        if value:
            line += f": {value}"
        print(line)
        dump_sr_tree(item, indent + 1)


def format_dataset_lines(ds, indent=0):
    """Recursively format every element in a dataset, full text values but
    binary VRs summarized as "<binary data, N bytes>" instead of dumped raw."""
    lines = []
    prefix = "  " * indent
    for elem in ds:
        tag_str = f"({elem.tag.group:04x},{elem.tag.element:04x})"
        name = elem.name
        vr = elem.VR or ""

        if vr == "SQ":
            items = elem.value or []
            lines.append(f"{prefix}{tag_str} {name}  SQ: {len(items)} item(s)")
            for i, item in enumerate(items):
                lines.append(f"{prefix}  [item {i}]")
                lines.extend(format_dataset_lines(item, indent + 2))
            continue

        if vr in BINARY_VRS:
            try:
                length = len(elem.value) if elem.value is not None else 0
            except TypeError:
                length = 0
            lines.append(f"{prefix}{tag_str} {name}  {vr}: <binary data, {length} bytes>")
            continue

        lines.append(f"{prefix}{tag_str} {name}  {vr}: {elem.value!r}")
    return lines


def print_raw_dataset(ds):
    print("\n  --- Full raw dataset (every tag pydicom recognizes, nested sequences included; "
          "binary fields summarized, not dumped) ---")
    for line in format_dataset_lines(ds):
        print(f"  {line}")


def print_sr_contents(ds):
    """Print a DICOM Structured Report's content tree (it has no pixel data
    to render - the report *is* this nested tree of text/coded findings),
    followed by the complete raw dataset so nothing is left out."""
    title = ""
    if ds.get("ConceptNameCodeSequence"):
        title = ds.ConceptNameCodeSequence[0].get("CodeMeaning", "")
    print(f"  SR Document Title: {title or '(none)'}")
    print(f"  CompletionFlag: {ds.get('CompletionFlag', '')}  "
          f"VerificationFlag: {ds.get('VerificationFlag', '')}")
    print("  Content (readable summary):")
    dump_sr_tree(ds)
    print_raw_dataset(ds)


def normalize_pixel_array(arr):
    """Reduce a pixel array of any DICOM shape down to (H, W) or (H, W, 3/4).

    Real-world exports (especially multi-frame ultrasound/color Doppler) can
    carry extra leading axes - e.g. (NumberOfFrames, Rows, Columns, Samples)
    even when NumberOfFrames=1, or stray singleton axes. Strip those down to
    a single displayable frame instead of handing PIL a shape it can't use.
    """
    arr = np.asarray(arr)
    is_color = arr.ndim >= 1 and arr.shape[-1] in (3, 4)
    target_ndim = 3 if is_color else 2

    while arr.ndim > target_ndim:
        arr = arr[arr.shape[0] // 2]  # collapse the outermost (frame-like) axis

    return arr


def dicom_to_png(dcm_path, png_path):
    try:
        ds = pydicom.dcmread(str(dcm_path), force=True)

        print(f"\n{dcm_path}")
        for tag in TAGS_TO_PRINT:
            value = ds.get(tag, "")
            if value != "":
                print(f"  {tag}: {value}")

        if "PixelData" not in ds:
            print("  (no pixel data - this isn't a renderable image; printing full contents instead)")
            if "ContentSequence" in ds:
                print_sr_contents(ds)
            else:
                print_raw_dataset(ds)
            return False

        try:
            arr = ds.pixel_array
        except Exception as exc:
            print(f"  (failed to decode pixel data: {exc})")
            print("  If this is a compressed transfer syntax, you may need an extra "
                  "codec package (e.g. pylibjpeg-openjpeg, python-gdcm).")
            return False

        print(f"  raw pixel_array shape={arr.shape} dtype={arr.dtype}")

        try:
            arr = pydicom.pixel_data_handlers.util.apply_voi_lut(arr, ds)
        except Exception:
            pass

        arr = normalize_pixel_array(arr)

        if arr.ndim not in (2, 3) or (arr.ndim == 3 and arr.shape[-1] not in (3, 4)):
            print(f"  (unexpected pixel array shape after normalization: {arr.shape} - skipping)")
            return False

        is_color = arr.ndim == 3

        if not is_color:
            arr = arr.astype(np.float32)
            lo, hi = np.percentile(arr, 0.5), np.percentile(arr, 99.5)
            if hi <= lo:
                lo, hi = arr.min(), arr.max()
            if hi <= lo:
                hi = lo + 1
            arr = np.clip((arr - lo) / (hi - lo), 0, 1) * 255
            arr = arr.astype(np.uint8)
            if getattr(ds, "PhotometricInterpretation", "") == "MONOCHROME1":
                arr = 255 - arr
            mode = "L"
        else:
            arr = arr.astype(np.uint8)
            mode = "RGB" if arr.shape[-1] == 3 else "RGBA"

        img = Image.fromarray(arr, mode=mode)
        img.save(png_path)
        print(f"  -> saved {png_path}  (final shape={arr.shape})")
        return True

    except Exception as exc:
        print(f"  (unexpected error on {dcm_path}: {type(exc).__name__}: {exc} - skipping)")
        return False


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    src = Path(sys.argv[1]).expanduser().resolve()

    if src.is_file():
        out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_suffix(".png")
        dicom_to_png(src, out)
        return

    if src.is_dir():
        out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else src / "png_preview"
        out_dir.mkdir(parents=True, exist_ok=True)
        converted = 0
        for path in sorted(src.rglob("*")):
            if not path.is_file():
                continue
            out_path = out_dir / (path.stem + ".png")
            if dicom_to_png(path, out_path):
                converted += 1
        print(f"\nConverted {converted} file(s) to: {out_dir}")
        return

    print(f"Not a file or folder: {src}")
    sys.exit(1)


if __name__ == "__main__":
    main()
