"""Load a Parquet file into an existing Postgres table, whole and unaltered.

    python parquet_to_postgres.py data.parquet --table measurements
    python parquet_to_postgres.py .                     # every *.parquet here
    python parquet_to_postgres.py data.parquet --truncate

The table is expected to exist already with the same columns as the file;
nothing is created, altered or dropped. Columns are matched by name, not by
position, so the two may be ordered differently. A column the table has and
the file does not is left to its default - a column the file has and the
table does not is an error, because that data would be silently discarded.

Connection details come from a .env file rather than the command line, so no
password ends up in shell history. Copy .env.example to .env and fill it in;
the script looks for .env beside itself, then in the current directory. Real
environment variables win over the file, which is what you want on a server.

Rows go in through COPY, in batches, inside one transaction: the whole file
lands or none of it does, and memory stays flat whatever the file's size.
Values are passed as CSV with every field quoted, so an empty string arrives
as an empty string and only a genuine null arrives as NULL - the two are not
the same thing and a plain to_csv dump confuses them.

Options:
    --table NAME       Target table, or schema.table. Default: $PGTABLE.
    --schema NAME      Schema, when not given as part of --table.
                       Default: $PGSCHEMA, else public.
    --env-file PATH    Read connection settings from here instead.
    --truncate         Empty the table first, in the same transaction.
    --batch-size N     Rows per COPY (default 50,000).
    --limit N          Stop after N rows in total, across files.
    --columns LIST     Only these columns, comma-separated.
    --allow-extra-file-columns
                       Skip file columns the table does not have instead of
                       refusing. Say so deliberately: it drops data.
    --dry-run          Read, check and convert, but write nothing.
"""

import argparse
import datetime as dt
import decimal
import io
import json
import math
import os
import sys

try:
    import pyarrow.parquet as pq
except ImportError:
    sys.exit("reading Parquet needs pyarrow:  pip install pyarrow")

try:
    import psycopg2
    from psycopg2 import sql
except ImportError:
    sys.exit("talking to Postgres needs psycopg2:  pip install psycopg2-binary")

DEFAULT_BATCH = 50_000
ENV_KEYS = ("DATABASE_URL", "PGHOST", "PGPORT", "PGDATABASE", "PGUSER",
            "PGPASSWORD", "PGSCHEMA", "PGTABLE", "PGSSLMODE")


def load_env(explicit=None):
    if explicit:
        paths = [explicit]
    else:
        here = os.path.dirname(os.path.abspath(__file__))
        paths = [os.path.join(here, ".env"), os.path.join(os.getcwd(), ".env")]

    for path in paths:
        if not os.path.exists(path):
            if explicit:
                sys.exit(f"no such file: {path}")
            continue
        with open(path, encoding="utf-8") as handle:
            for number, line in enumerate(handle, start=1):
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("export "):
                    line = line[len("export "):].lstrip()
                if "=" not in line:
                    sys.exit(f"{path}:{number}: expected KEY=VALUE")
                key, value = line.split("=", 1)
                key, value = key.strip(), value.strip()
                if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
                    value = value[1:-1]
                os.environ.setdefault(key, value)
        return path
    return None


def connect():
    url = os.environ.get("DATABASE_URL", "").strip()
    try:
        if url:
            return psycopg2.connect(url)
        missing = [k for k in ("PGHOST", "PGDATABASE", "PGUSER")
                   if not os.environ.get(k, "").strip()]
        if missing:
            sys.exit("connection settings are incomplete: "
                     + ", ".join(missing) + " not set.\n"
                     "  Copy .env.example to .env and fill it in, or set"
                     " DATABASE_URL.")
        return psycopg2.connect(
            host=os.environ["PGHOST"],
            port=os.environ.get("PGPORT", "5432"),
            dbname=os.environ["PGDATABASE"],
            user=os.environ["PGUSER"],
            password=os.environ.get("PGPASSWORD") or None,
            sslmode=os.environ.get("PGSSLMODE") or None,
            connect_timeout=int(os.environ.get("PGCONNECT_TIMEOUT", "10")),
        )
    except psycopg2.Error as exc:
        sys.exit(f"could not connect: {str(exc).strip()}")


def scalar_text(value):
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return f"{value.total_seconds()} seconds"
    if isinstance(value, decimal.Decimal):
        return str(value)
    if isinstance(value, (bytes, bytearray, memoryview)):
        return "\\x" + bytes(value).hex()
    if isinstance(value, float):
        if math.isnan(value):
            return "NaN"
        if math.isinf(value):
            return "Infinity" if value > 0 else "-Infinity"
        return repr(value)
    return str(value)


def array_literal(value):
    if value is None:
        return "NULL"
    if isinstance(value, (list, tuple)):
        return "{" + ",".join(array_literal(v) for v in value) + "}"
    text = scalar_text(value)
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def encode(value, is_array):
    if value is None:
        return None
    if isinstance(value, (list, tuple)):
        return array_literal(value) if is_array else json_text(value)
    if isinstance(value, dict):
        return json_text(value)
    return scalar_text(value)


def json_text(value):
    return json.dumps(value, default=scalar_text, ensure_ascii=False)


def split_table(name, default_schema):
    if name and "." in name:
        schema, _, table = name.partition(".")
        return schema.strip('"'), table.strip('"')
    return default_schema, name


def table_columns(cursor, schema, table):
    cursor.execute("""
        select column_name, udt_name
        from information_schema.columns
        where table_schema = %s and table_name = %s
        order by ordinal_position
    """, (schema, table))
    return dict(cursor.fetchall())


def check_columns(file_columns, table_cols, schema, table, allow_extra):
    if not table_cols:
        sys.exit(f'table "{schema}"."{table}" does not exist, or this user'
                 " cannot see it.\n  This script loads into an existing table;"
                 " it does not create one.")

    unknown = [c for c in file_columns if c not in table_cols]
    if unknown and not allow_extra:
        sys.exit(f'these columns are in the file but not in "{schema}"."'
                 f'{table}":\n    ' + ", ".join(unknown)
                 + "\n  Add them to the table, select columns with --columns,"
                 "\n  or pass --allow-extra-file-columns to drop them.")

    used = [c for c in file_columns if c in table_cols]
    if not used:
        sys.exit("no column in the file matches a column in the table.")
    absent = [c for c in table_cols if c not in used]
    return used, unknown, absent


def csv_field(text):
    if text is None:
        return ""
    return '"' + text.replace('"', '""') + '"'


def batch_to_csv(batch, columns, array_flags):
    data = [batch.column(batch.schema.get_field_index(c)).to_pylist()
            for c in columns]
    lines = []
    for row in zip(*data):
        lines.append(",".join(
            csv_field(encode(value, array_flags[i]))
            for i, value in enumerate(row)))
    lines.append("")
    return io.StringIO("\n".join(lines))


def copy_file(cursor, path, schema, table, columns, array_flags,
              batch_size, limit):
    statement = sql.SQL("COPY {}.{} ({}) FROM STDIN WITH (FORMAT csv,"
                        " NULL '', QUOTE '\"')").format(
        sql.Identifier(schema), sql.Identifier(table),
        sql.SQL(", ").join(sql.Identifier(c) for c in columns))

    sent = 0
    batches = 0
    reader = pq.ParquetFile(path)
    for batch in reader.iter_batches(batch_size=batch_size, columns=columns):
        if limit is not None and sent + batch.num_rows > limit:
            batch = batch.slice(0, limit - sent)
        if batch.num_rows == 0:
            break
        cursor.copy_expert(statement, batch_to_csv(batch, columns,
                                                   array_flags))
        sent += batch.num_rows
        batches += 1
        if batches > 1:
            print(f"    {sent:,} rows", end="\r", flush=True)
        if limit is not None and sent >= limit:
            break
    return sent


def parquet_files(target):
    if os.path.isdir(target):
        found = sorted(os.path.join(target, n) for n in os.listdir(target)
                       if n.lower().endswith(".parquet"))
        if not found:
            sys.exit(f"no .parquet files in {target}")
        return found
    if not os.path.exists(target):
        sys.exit(f"no such file: {target}")
    return [target]


def file_schema(paths, wanted):
    first = [f.name for f in pq.ParquetFile(paths[0]).schema_arrow]
    for path in paths[1:]:
        other = {f.name for f in pq.ParquetFile(path).schema_arrow}
        if set(first) != other:
            sys.exit(f"{path} has different columns from {paths[0]}.\n"
                     "  Load them separately, or use --columns for the"
                     " columns they share.")
    if wanted:
        missing = [c for c in wanted if c not in first]
        if missing:
            sys.exit("--columns names columns the file does not have: "
                     + ", ".join(missing))
        return list(wanted)
    return first


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("parquet", help="a .parquet file, or a folder of them")
    ap.add_argument("--table", default=None)
    ap.add_argument("--schema", default=None)
    ap.add_argument("--env-file", default=None)
    ap.add_argument("--truncate", action="store_true")
    ap.add_argument("--batch-size", type=int, default=DEFAULT_BATCH)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--columns", default=None)
    ap.add_argument("--allow-extra-file-columns", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    env_path = load_env(args.env_file)
    table_name = args.table or os.environ.get("PGTABLE", "").strip()
    if not table_name:
        sys.exit("no target table: pass --table, or set PGTABLE in .env")
    schema, table = split_table(
        table_name, args.schema or os.environ.get("PGSCHEMA") or "public")

    paths = parquet_files(args.parquet)
    wanted = ([c.strip() for c in args.columns.split(",") if c.strip()]
              if args.columns else None)
    columns = file_schema(paths, wanted)

    if env_path:
        print(f"settings from {env_path}")
    connection = connect()
    connection.autocommit = False
    try:
        with connection.cursor() as cursor:
            table_cols = table_columns(cursor, schema, table)
            columns, dropped, absent = check_columns(
                columns, table_cols, schema, table,
                args.allow_extra_file_columns)
            array_flags = [table_cols[c].startswith("_") for c in columns]

            print(f'into "{schema}"."{table}": {len(columns)} columns'
                  f' from {len(paths)} file' + ("s" if len(paths) > 1 else ""))
            for column in dropped:
                print(f"  dropping {column}: no such column in the table")
            for column in absent:
                print(f"  {column}: not in the file, left to its default")

            if args.dry_run:
                for path in paths:
                    for batch in pq.ParquetFile(path).iter_batches(
                            batch_size=min(args.batch_size, 1000),
                            columns=columns):
                        batch_to_csv(batch, columns, array_flags)
                        break
                print("dry run: nothing was written")
                connection.rollback()
                return

            if args.truncate:
                cursor.execute(sql.SQL("TRUNCATE {}.{}").format(
                    sql.Identifier(schema), sql.Identifier(table)))
                print("  table emptied (rolled back if anything below fails)")

            total = 0
            for path in paths:
                print(f"  {os.path.basename(path)}")
                remaining = (None if args.limit is None
                             else max(args.limit - total, 0))
                rows = copy_file(cursor, path, schema, table, columns,
                                 array_flags, args.batch_size,
                                 remaining)
                total += rows
                print(f"    {rows:,} rows            ")
                if args.limit is not None and total >= args.limit:
                    break
        connection.commit()
        print(f'committed {total:,} rows into "{schema}"."{table}"')
    except psycopg2.Error as exc:
        connection.rollback()
        sys.exit(f"\nnothing was written - rolled back.\n  "
                 + str(exc).strip())
    except KeyboardInterrupt:
        connection.rollback()
        sys.exit("\ninterrupted - rolled back, the table is as it was.")
    finally:
        connection.close()


if __name__ == "__main__":
    main()
