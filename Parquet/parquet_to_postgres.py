"""Wipe a Postgres table and load a Parquet file into it.

    python parquet_to_postgres.py data.parquet

The table is emptied and the whole file is loaded in one transaction, so a
failure anywhere leaves the old rows untouched. Columns are matched by name,
so the file and the table may order them differently. Connection settings
and the target table come from .env - copy .env.example and fill it in.
"""

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

BATCH_ROWS = 50_000


def load_env():
    here = os.path.dirname(os.path.abspath(__file__))
    for path in (os.path.join(here, ".env"),
                 os.path.join(os.getcwd(), ".env")):
        if not os.path.exists(path):
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


def target_table():
    name = os.environ.get("PGTABLE", "").strip()
    if not name:
        sys.exit("no target table: set PGTABLE in .env")
    if "." in name:
        schema, _, table = name.partition(".")
        return schema.strip('"'), table.strip('"')
    return (os.environ.get("PGSCHEMA") or "public"), name


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


def encode(value):
    if value is None:
        return None
    if isinstance(value, (list, tuple)):
        return array_literal(value)
    if isinstance(value, dict):
        return json.dumps(value, default=scalar_text, ensure_ascii=False)
    return scalar_text(value)


def csv_field(text):
    if text is None:
        return ""
    return '"' + text.replace('"', '""') + '"'


def batch_to_csv(batch):
    data = [column.to_pylist() for column in batch.columns]
    lines = [",".join(csv_field(encode(value)) for value in row)
             for row in zip(*data)]
    lines.append("")
    return io.StringIO("\n".join(lines))


def main():
    if len(sys.argv) != 2 or sys.argv[1].startswith("-"):
        sys.exit("usage: python parquet_to_postgres.py <file.parquet>")
    path = sys.argv[1]
    if not os.path.exists(path):
        sys.exit(f"no such file: {path}")

    env_path = load_env()
    schema, table = target_table()
    reader = pq.ParquetFile(path)
    columns = [field.name for field in reader.schema_arrow]

    if env_path:
        print(f"settings from {env_path}")
    connection = connect()
    connection.autocommit = False
    statement = sql.SQL("COPY {}.{} ({}) FROM STDIN WITH (FORMAT csv,"
                        " NULL '', QUOTE '\"')").format(
        sql.Identifier(schema), sql.Identifier(table),
        sql.SQL(", ").join(sql.Identifier(c) for c in columns))
    try:
        with connection.cursor() as cursor:
            cursor.execute(sql.SQL("TRUNCATE {}.{}").format(
                sql.Identifier(schema), sql.Identifier(table)))
            rows = 0
            for batch in reader.iter_batches(batch_size=BATCH_ROWS):
                cursor.copy_expert(statement, batch_to_csv(batch))
                rows += batch.num_rows
        connection.commit()
        print(f'{rows:,} rows into "{schema}"."{table}"')
    except psycopg2.Error as exc:
        connection.rollback()
        sys.exit(f"nothing was written - rolled back, the table is as it"
                 f" was.\n  {str(exc).strip()}")
    finally:
        connection.close()


if __name__ == "__main__":
    main()
