"""Convert a Parquet file to Excel or CSV, with timestamps that read correctly.

    python parquet_to_excel.py data.parquet                 # -> data.xlsx
    python parquet_to_excel.py data.parquet --csv           # -> data.csv
    python parquet_to_excel.py data.parquet -o report.xlsx

Timestamps are the part that goes wrong. Parquet stores them as
datetime64[us] - microseconds since the epoch - and left alone they arrive in
Excel as a number like 45678.375, or in CSV as 1710000000000000. Three things
are needed to avoid that, and this does all three:

  * every timestamp column is given an Excel number format, so the cell shows
    a date rather than the serial number behind it;
  * timezone-aware columns are converted (by default to UTC, see --tz) and the
    offset dropped, because openpyxl refuses to write an aware datetime and
    Excel has nowhere to keep one;
  * microseconds survive to CSV exactly; Excel keeps time as a fraction of a
    day, so sub-millisecond precision is not representable there. When a
    column actually uses microseconds this says so rather than rounding
    silently - use --csv when those digits matter.

Options:
    -o, --out PATH    Output file. Default: the input name with a new suffix.
    --csv             Write CSV instead of Excel.
    --sheet NAME      Worksheet name (default "Data").
    --tz {utc,local,keep-naive}
                      What to do with timezone-aware timestamps. utc (default)
                      converts to UTC; local converts to this machine's zone;
                      keep-naive drops the zone without converting. All three
                      then write a naive datetime, which is all Excel holds.
    --datetime-format F
                      Excel display format (default "yyyy-mm-dd hh:mm:ss") or,
                      with --csv, a strftime pattern.
    --columns LIST    Only these columns, comma-separated, in this order.
    --limit N         Only the first N rows.
    --sheet-per-chunk Excel caps at 1,048,575 data rows. By default a bigger
                      table is refused; this splits it across sheets instead.
"""

import argparse
import os
import sys

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas is required:  pip install pandas")

EXCEL_MAX_ROWS = 1_048_576          # including the header row
DEFAULT_EXCEL_FORMAT = "yyyy-mm-dd hh:mm:ss"
DEFAULT_CSV_FORMAT = "%Y-%m-%d %H:%M:%S.%f"


def read_parquet(path, columns):
    try:
        return pd.read_parquet(path, columns=columns)
    except ImportError:
        sys.exit("reading Parquet needs an engine:  pip install pyarrow")
    except Exception as exc:
        sys.exit(f"could not read {path}: {type(exc).__name__}: {exc}")


def timestamp_columns(df):
    """Columns holding timestamps, tz-aware or not.

    Checked by dtype kind rather than by name: a column called "date" may be
    text, and a column called "x" may be a timestamp.
    """
    return [c for c in df.columns
            if pd.api.types.is_datetime64_any_dtype(df[c])]


def strip_timezone(df, columns, mode):
    """Excel has no concept of an offset and openpyxl refuses to write one."""
    converted = []
    for col in columns:
        tz = getattr(df[col].dtype, "tz", None)
        if tz is None:
            continue
        if mode == "utc":
            df[col] = df[col].dt.tz_convert("UTC").dt.tz_localize(None)
        elif mode == "local":
            df[col] = df[col].dt.tz_convert(
                pd.Timestamp.now().astimezone().tzinfo).dt.tz_localize(None)
        else:
            df[col] = df[col].dt.tz_localize(None)
        converted.append((col, str(tz)))
    return converted


def uses_microseconds(series):
    """True when any value has sub-millisecond detail Excel cannot hold."""
    values = series.dropna()
    if values.empty:
        return False
    try:
        return bool((values.dt.microsecond % 1000 != 0).any())
    except (AttributeError, TypeError):
        return False


def write_csv(df, out, fmt):
    df.to_csv(out, index=False, date_format=fmt)


def write_excel(df, out, sheet, fmt, split):
    from openpyxl.utils import get_column_letter

    rows = len(df)
    per_sheet = EXCEL_MAX_ROWS - 1
    if rows > per_sheet and not split:
        sys.exit(f"{rows:,} rows exceeds Excel's limit of {per_sheet:,}.\n"
                 "  --csv writes it in one piece, or --sheet-per-chunk splits "
                 "it across sheets.")

    stamps = timestamp_columns(df)
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format=fmt) as writer:
        chunks = [(sheet, df)] if rows <= per_sheet else [
            (f"{sheet}_{i + 1}", df.iloc[i * per_sheet:(i + 1) * per_sheet])
            for i in range((rows + per_sheet - 1) // per_sheet)]
        for name, chunk in chunks:
            chunk.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            ws.freeze_panes = "A2"
            for idx, column in enumerate(chunk.columns, start=1):
                letter = get_column_letter(idx)
                # A date in a General cell shows as its serial number, so the
                # format goes on every cell of the column, not just the header.
                if column in stamps:
                    for cell in ws[letter][1:]:
                        cell.number_format = fmt
                width = max(len(str(column)) + 2, len(fmt) + 2 if column in
                            stamps else 12)
                ws.column_dimensions[letter].width = min(width, 40)
    return len(chunks) if rows > per_sheet else 1


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("parquet")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--csv", action="store_true")
    ap.add_argument("--sheet", default="Data")
    ap.add_argument("--tz", default="utc",
                    choices=["utc", "local", "keep-naive"])
    ap.add_argument("--datetime-format", default=None)
    ap.add_argument("--columns", default=None)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--sheet-per-chunk", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.parquet):
        sys.exit(f"no such file: {args.parquet}")
    columns = ([c.strip() for c in args.columns.split(",") if c.strip()]
               if args.columns else None)
    df = read_parquet(args.parquet, columns)
    if args.limit is not None:
        df = df.head(args.limit)

    out = args.out or os.path.splitext(args.parquet)[0] + (
        ".csv" if args.csv else ".xlsx")
    fmt = args.datetime_format or (
        DEFAULT_CSV_FORMAT if args.csv else DEFAULT_EXCEL_FORMAT)

    stamps = timestamp_columns(df)
    precise = [c for c in stamps if uses_microseconds(df[c])]
    converted = strip_timezone(df, stamps, args.tz)

    if args.csv:
        write_csv(df, out, fmt)
        sheets = 0
    else:
        sheets = write_excel(df, out, args.sheet, fmt, args.sheet_per_chunk)

    print(f"wrote {out}")
    print(f"  {len(df):,} rows x {len(df.columns)} columns"
          + (f" across {sheets} sheets" if sheets > 1 else ""))
    if stamps:
        print(f"  timestamp columns: {', '.join(map(str, stamps))}")
        print(f"  shown as: {fmt}")
    for col, tz in converted:
        print(f"  {col}: converted from {tz} and the offset dropped"
              f" ({args.tz})")
    if precise and not args.csv:
        print(f"  NOTE: {', '.join(map(str, precise))} use microseconds, which"
              " Excel cannot\n        represent - those digits are rounded."
              " Use --csv to keep them exactly.")


if __name__ == "__main__":
    main()
