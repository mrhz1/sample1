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


TAGS_TO_PRINT = [
    "Modality", "StudyDescription", "SeriesDescription", "BodyPartExamined",
    "StudyDate", "SeriesNumber", "InstanceNumber", "Rows", "Columns",
    "NumberOfFrames", "SamplesPerPixel", "PhotometricInterpretation",
    "PlanarConfiguration",
]


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
            print("  (no pixel data - can't render an image)")
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
