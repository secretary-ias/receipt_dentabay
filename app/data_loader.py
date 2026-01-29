"""Utilities to transform MySQL data into a lean SQLite cache."""
from __future__ import annotations

import csv
import hashlib
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional, Sequence

try:
    import pymysql
except ImportError:  # pragma: no cover - handled at runtime
    pymysql = None  # type: ignore


class DataImportError(Exception):
    """Raised when the data source cannot be processed."""


def _to_none_or_str(value: Optional[str]) -> Optional[str]:
    if value is None:
        return None
    return value.replace("\r\n", "\n")


def _to_clean_str(value: Optional[str]) -> str:
    if value is None:
        return ""
    return value.replace("\r\n", "\n").strip()


def _to_float(value: Optional[str]) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except ValueError:
        return 0.0


def _to_int(value: Optional[str]) -> int:
    if value is None or value == "":
        return 0
    try:
        return int(float(value))
    except ValueError:
        return 0


@dataclass(frozen=True)
class TableImportPlan:
    table: str
    create_sql: str
    insert_sql: str
    index_sql: Sequence[str]
    index_map: Sequence[int]
    converters: Sequence[Callable[[Optional[str]], object]]
    select_columns: Sequence[str]


class ClinicDataImporter:
    """Import DoctorAssist data (backup SQL or live MySQL) into SQLite."""

    ENCODING = "utf-8"

    def __init__(
        self,
        sqlite_path: Path,
        backup_path: Optional[Path] = None,
        mysql_settings: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.sqlite_path = Path(sqlite_path)
        self.backup_path = Path(backup_path) if backup_path else None
        self.mysql_settings = mysql_settings

        if self.backup_path and not self.backup_path.exists():
            raise DataImportError(f"Backup file not found: {self.backup_path}")
        if mysql_settings and pymysql is None:
            raise DataImportError("pymysql is required for live MySQL imports. Please install it and retry.")
        if not self.backup_path and not self.mysql_settings:
            raise DataImportError("Either a backup path or MySQL settings must be supplied.")

        self.mode = "backup" if self.backup_path else "mysql"

    def ensure_cache(self, force: bool = False, progress: Optional[Callable[[str], None]] = None) -> None:
        """Ensure the SQLite cache reflects the selected data source."""
        reporter = progress or (lambda _msg: None)

        signature = self._current_signature()
        if self.mode == "backup" and not force and self.sqlite_path.exists():
            try:
                with sqlite3.connect(self.sqlite_path) as conn:
                    conn.row_factory = sqlite3.Row
                    cursor = conn.execute("SELECT value FROM metadata WHERE key = 'source_signature'")
                    row = cursor.fetchone()
                    if not row:
                        cursor = conn.execute("SELECT value FROM metadata WHERE key = 'backup_sha256'")
                        row = cursor.fetchone()
                    if row and row[0] == signature:
                        reporter("SQLite cache already up to date.")
                        return
            except sqlite3.Error:
                pass

        reporter("Rebuilding SQLite cache ...")
        self._rebuild_sqlite(signature, reporter)

    def _current_signature(self) -> str:
        if self.mode == "backup":
            assert self.backup_path is not None
            sha256 = hashlib.sha256()
            with self.backup_path.open("rb") as fh:
                for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                    sha256.update(chunk)
            return sha256.hexdigest()
        cfg = self.mysql_settings or {}
        return "mysql://{host}:{port}/{database}".format(
            host=cfg.get("host", "localhost"),
            port=cfg.get("port", 3306),
            database=cfg.get("database", "clinicdb"),
        )

    def _rebuild_sqlite(self, signature: str, progress: Callable[[str], None]) -> None:
        """Recreate the SQLite cache in place."""
        self.sqlite_path.parent.mkdir(parents=True, exist_ok=True)

        mysql_conn = None
        try:
            if self.mode == "mysql":
                mysql_conn = self._open_mysql()

            with sqlite3.connect(self.sqlite_path, timeout=5) as conn:
                conn.execute("PRAGMA journal_mode = OFF;")
                conn.execute("PRAGMA synchronous = OFF;")
                conn.execute("PRAGMA temp_store = MEMORY;")

                for plan in self._plans():
                    conn.execute(f"DROP TABLE IF EXISTS {plan.table}")
                    conn.execute(plan.create_sql)

                conn.execute("CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL)")
                conn.commit()

                for plan in self._plans():
                    self._import_table(conn, plan, progress, mysql_conn)

                conn.execute(
                    "INSERT OR REPLACE INTO metadata(key, value) VALUES('source_signature', ?)",
                    (signature,),
                )
                conn.commit()
        finally:
            if mysql_conn is not None:
                mysql_conn.close()
        progress("SQLite cache rebuild complete.")


    def _open_mysql(self):  # type: ignore[override]
        assert self.mysql_settings is not None

        def _connect(charset: str):
            return pymysql.connect(
                host=self.mysql_settings.get("host", "localhost"),
                port=int(self.mysql_settings.get("port", 3306)),
                user=self.mysql_settings.get("user", ""),
                password=self.mysql_settings.get("password", ""),
                database=self.mysql_settings.get("database", "clinicdb"),
                charset=charset,
                cursorclass=pymysql.cursors.Cursor,
            )

        try:
            return _connect("utf8mb4")
        except pymysql.err.OperationalError as exc:  # type: ignore[attr-defined]
            error_code = exc.args[0] if exc.args else None
            if error_code == 1115 or "Unknown character set" in str(exc):
                # Older MySQL releases (5.5.2 and earlier) lack utf8mb4 support; retry with utf8.
                try:
                    return _connect("utf8")
                except pymysql.MySQLError as fallback_exc:  # type: ignore[attr-defined]
                    raise DataImportError(
                        "Unable to connect to MySQL using utf8mb4; fallback to utf8 failed: {exc}".format(
                            exc=fallback_exc,
                        )
                    ) from fallback_exc
            raise DataImportError(f"Unable to connect to MySQL: {exc}") from exc
        except pymysql.MySQLError as exc:  # type: ignore[attr-defined]
            raise DataImportError(f"Unable to connect to MySQL: {exc}") from exc

    def _import_table(
        self,
        conn: sqlite3.Connection,
        plan: TableImportPlan,
        progress: Callable[[str], None],
        mysql_conn,
    ) -> None:
        progress(f"Importing {plan.table} ...")
        batch: List[tuple] = []
        total = 0
        for rows in self._iter_table_rows(plan, mysql_conn):
            for raw_row in rows:
                prepared: List[object] = []
                for source_index, converter in zip(plan.index_map, plan.converters):
                    token = raw_row[source_index] if source_index < len(raw_row) else None
                    prepared.append(converter(self._normalise_token(token)))
                batch.append(tuple(prepared))
            if len(batch) >= 500:
                conn.executemany(plan.insert_sql, batch)
                total += len(batch)
                batch.clear()
                progress(f"{plan.table}: {total} rows processed...")
        if batch:
            conn.executemany(plan.insert_sql, batch)
            total += len(batch)
            batch.clear()
        for index_sql in plan.index_sql:
            conn.execute(index_sql)
        conn.commit()
        progress(f"{plan.table}: {total} rows imported.")

    def _normalise_token(self, token: Optional[str]) -> Optional[str]:
        if token is None:
            return None
        value = token.strip()
        if value.upper() == "NULL":
            return None
        return value

    def _iter_table_rows(
        self,
        plan: TableImportPlan,
        mysql_conn,
    ) -> Iterator[List[List[Optional[str]]]]:
        if self.mode == "backup":
            assert self.backup_path is not None
            yield from self._iter_table_rows_backup(plan.table)
        else:
            assert mysql_conn is not None
            yield from self._iter_table_rows_mysql(plan, mysql_conn)

    def _iter_table_rows_backup(self, table: str) -> Iterator[List[List[Optional[str]]]]:
        target = f"INSERT INTO `{table}` VALUES"
        buffer: Optional[str] = None
        with self.backup_path.open("r", encoding=self.ENCODING, errors="ignore") as fh:  # type: ignore[arg-type]
            for line in fh:
                stripped = line.strip()
                if not stripped:
                    continue
                if buffer is not None:
                    buffer += stripped
                    if stripped.endswith(";"):
                        yield self._parse_insert_values(buffer, target)
                        buffer = None
                elif stripped.startswith(target):
                    buffer = stripped
                    if stripped.endswith(";"):
                        yield self._parse_insert_values(buffer, target)
                        buffer = None
        if buffer:
            yield self._parse_insert_values(buffer, target)

    def _iter_table_rows_mysql(
        self,
        plan: TableImportPlan,
        mysql_conn,
    ) -> Iterator[List[List[Optional[str]]]]:
        columns = ", ".join(f"`{col}`" for col in plan.select_columns)
        sql = f"SELECT {columns} FROM `{plan.table}`"
        with mysql_conn.cursor() as cursor:
            cursor.execute(sql)
            while True:
                rows = cursor.fetchmany(500)
                if not rows:
                    break
                formatted: List[List[Optional[str]]] = []
                for row in rows:
                    prepared: List[Optional[str]] = []
                    for value in row:
                        if value is None:
                            prepared.append(None)
                        elif isinstance(value, bytes):
                            prepared.append(value.decode("utf-8", errors="ignore"))
                        else:
                            prepared.append(str(value))
                    formatted.append(prepared)
                yield formatted

    def _parse_insert_values(self, statement: str, prefix: str) -> List[List[Optional[str]]]:
        if not statement.startswith(prefix):
            raise DataImportError(f"Unexpected INSERT statement: {statement[:80]}...")
        payload = statement[len(prefix):].lstrip()
        if payload.startswith("VALUES"):
            body = payload[len("VALUES"):]
        else:
            body = payload
        body = body.rstrip(";")

        rows: List[str] = []
        current = ""
        depth = 0
        for ch in body:
            current += ch
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    rows.append(current.strip())
                    current = ""
        parsed_rows: List[List[Optional[str]]] = []
        for row in rows:
            fragment = row
            if fragment.startswith("(") and fragment.endswith(")"):
                fragment = fragment[1:-1]
            reader = csv.reader([fragment], delimiter=",", quotechar="'", escapechar="\\")
            parsed_rows.append(list(next(reader)))
        return parsed_rows
    def _plans(self) -> Sequence[TableImportPlan]:
        return (
            TableImportPlan(
                table="patients",
                create_sql=(
                    "CREATE TABLE patients ("
                    " icpassport TEXT PRIMARY KEY,"
                    " name TEXT NOT NULL,"
                    " receipt_name TEXT,"
                    " preferred_name TEXT,"
                    " company TEXT,"
                    " phone_fixed TEXT,"
                    " phone_mobile TEXT,"
                    " email TEXT"
                    ")"
                ),
                insert_sql="INSERT OR REPLACE INTO patients VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                index_sql=(
                    "CREATE INDEX IF NOT EXISTS idx_patients_name ON patients(name)",
                ),
                index_map=(1, 2, 21, 29, 16, 10, 11, 30),
                converters=(
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                ),
                select_columns=(
                    "register_date",
                    "icpassport",
                    "name",
                    "sex",
                    "DOB",
                    "address",
                    "city",
                    "state",
                    "country",
                    "postcode",
                    "phone_fixed",
                    "phone_mobile",
                    "remark",
                    "medical_illness",
                    "occupation",
                    "removed",
                    "company",
                    "photo",
                    "race",
                    "marketing",
                    "source",
                    "receipt_name",
                    "Title",
                    "PatientStatus",
                    "EmergencyContact",
                    "EmergencyPhoneNo",
                    "BillingType",
                    "companyaddress",
                    "companycontact",
                    "preferredname",
                    "Emailaddress",
                    "Memberid",
                    "app_doublevalue",
                    "username",
                    "modified_date",
                    "language",
                    "contact_relation",
                    "religion",
                    "pat_id",
                    "national",
                    "lastName",
                    "ref_id",
                    "priority_id",
                    "fingerprint",
                ),
            ),
            TableImportPlan(
                table="payment_method",
                create_sql="CREATE TABLE payment_method (paycode TEXT PRIMARY KEY, description TEXT NOT NULL)",
                insert_sql="INSERT OR REPLACE INTO payment_method VALUES (?, ?)",
                index_sql=(),
                index_map=(0, 1),
                converters=(
                    _to_clean_str,
                    _to_clean_str,
                ),
                select_columns=("paycode", "description"),
            ),
            TableImportPlan(
                table="stock_items",
                create_sql=(
                    "CREATE TABLE stock_items ("
                    " id TEXT PRIMARY KEY,"
                    " name TEXT NOT NULL,"
                    " category TEXT,"
                    " selling_price REAL,"
                    " removed INTEGER,"
                    " is_service INTEGER,"
                    " unit_cost REAL"
                    ")"
                ),
                insert_sql="INSERT OR REPLACE INTO stock_items VALUES (?, ?, ?, ?, ?, ?, ?)",
                index_sql=(
                    "CREATE INDEX IF NOT EXISTS idx_stock_items_name ON stock_items(name)",
                ),
                index_map=(0, 1, 3, 4, 5, 6, 23),
                converters=(
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_float,
                    _to_int,
                    _to_int,
                    _to_float,
                ),
                select_columns=(
                    "id",
                    "name",
                    "company",
                    "category",
                    "selling_price",
                    "removed",
                    "is_service",
                    "dose",
                    "times",
                    "day",
                    "procedure",
                    "is_print",
                    "proposed_qty",
                    "proposed_remark",
                    "doseqty",
                    "timesqty",
                    "stockcode",
                    "alternatecode",
                    "username",
                    "modified_date",
                    "restock_level",
                    "ratio",
                    "tax_code",
                    "unit_cost",
                    "commission",
                    "is_restrict",
                ),
            ),
            TableImportPlan(
                table="receipts",
                create_sql=(
                    "CREATE TABLE receipts ("
                    " rcpt_id TEXT PRIMARY KEY,"
                    " issued TEXT NOT NULL,"
                    " patient_id TEXT NOT NULL,"
                    " total REAL,"
                    " subtotal REAL,"
                    " gst REAL,"
                    " payment_code TEXT,"
                    " remark TEXT,"
                    " removed INTEGER,"
                    " discount REAL,"
                    " rounding REAL,"
                    " consult_fees REAL,"
                    " done_by TEXT,"
                    " department_type TEXT,"
                    " settled_by TEXT,"
                    " tax_total REAL"
                    ")"
                ),
                insert_sql="INSERT OR REPLACE INTO receipts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                index_sql=(
                    "CREATE INDEX IF NOT EXISTS idx_receipts_patient_date ON receipts(patient_id, issued)",
                ),
                index_map=(0, 1, 2, 4, 9, 10, 5, 7, 8, 16, 18, 3, 17, 20, 14, 15),
                converters=(
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_float,
                    _to_float,
                    _to_float,
                    _to_clean_str,
                    _to_clean_str,
                    _to_int,
                    _to_float,
                    _to_float,
                    _to_float,
                    _to_clean_str,
                    _to_clean_str,
                    _to_clean_str,
                    _to_float,
                ),
                select_columns=(
                    "rcpt_id",
                    "issued",
                    "patient_id",
                    "consult_fees",
                    "total",
                    "payment",
                    "Billed",
                    "Remark",
                    "removed",
                    "subtotal",
                    "gst",
                    "itemize",
                    "username",
                    "modified_date",
                    "settledby",
                    "tax_total",
                    "disc_total",
                    "done_by",
                    "rounding",
                    "mr_id",
                    "department_type",
                    "printed_name",
                ),
            ),
            TableImportPlan(
                table="receipt_items",
                create_sql=(
                    "CREATE TABLE receipt_items ("
                    " id INTEGER PRIMARY KEY,"
                    " rcpt_id TEXT NOT NULL,"
                    " item_id TEXT,"
                    " qty INTEGER,"
                    " unit_price REAL,"
                    " subtotal REAL,"
                    " discount REAL,"
                    " username TEXT,"
                    " remark TEXT"
                    ")"
                ),
                insert_sql="INSERT OR REPLACE INTO receipt_items VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                index_sql=(
                    "CREATE INDEX IF NOT EXISTS idx_receipt_items_rcpt ON receipt_items(rcpt_id)",
                    "CREATE INDEX IF NOT EXISTS idx_receipt_items_item ON receipt_items(item_id)",
                ),
                index_map=(0, 1, 2, 3, 4, 5, 18, 10, 8),
                converters=(
                    _to_int,
                    _to_clean_str,
                    _to_clean_str,
                    _to_int,
                    _to_float,
                    _to_float,
                    _to_float,
                    _to_clean_str,
                    _to_clean_str,
                ),
                select_columns=(
                    "ID",
                    "rcpt_id",
                    "item",
                    "qty",
                    "unitprice",
                    "subtotal",
                    "DoseID",
                    "TimesID",
                    "rcpt_remark",
                    "day",
                    "username",
                    "doseqty",
                    "timesqty",
                    "bodypart",
                    "pres_check",
                    "modified_date",
                    "tax_subtotal",
                    "disc_amount",
                    "unit_fixed_cost",
                    "commission",
                ),
            ),
        )
