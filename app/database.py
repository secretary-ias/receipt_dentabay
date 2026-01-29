"""Data access layer for the clinic receipt database."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Sequence, Tuple

import pymysql
from pymysql.cursors import DictCursor
from pymysql import err as pymysql_errors


@dataclass(frozen=True)
class Patient:
    icpassport: str
    name: str
    receipt_name: str
    preferred_name: str
    company: str
    phone_fixed: str
    phone_mobile: str
    email: str


@dataclass(frozen=True)
class Receipt:
    rcpt_id: str
    issued: datetime
    patient_id: str
    total: float
    subtotal: float
    gst: float
    payment_code: str
    remark: str
    discount: float
    rounding: float
    consult_fees: float
    done_by: str
    department_type: str
    settled_by: str
    tax_total: float
    mr_id: int


@dataclass(frozen=True)
class ReceiptItem:
    id: int
    item_id: str
    name: str
    qty: int
    unit_price: float
    subtotal: float
    discount: float
    username: str
    remark: str


@dataclass(frozen=True)
class PartialPayment:
    payment_id: int
    receipt_id: str
    date: datetime
    amount: float
    pay_code: str
    method: str
    username: str
    settled_by: str
    remark: str


@dataclass(frozen=True)
class ReceiptSummary:
    receipt: Receipt
    patient: Patient


@dataclass(frozen=True)
class AppointmentDetail:
    id: int
    scheduled: datetime
    patient_id: str
    patient_name: str
    reason: str
    resource: str
    appointment_type: str
    location: str
    queue_number: str
    created_by: str
    status: str
    status_color: str
    status_id: int


@dataclass(frozen=True)
class PatientProfile:
    patient_id: str
    name: str
    preferred_name: str
    receipt_name: str
    sex: str
    date_of_birth: date
    phone_mobile: str
    phone_fixed: str
    email: str
    address: str
    city: str
    state: str
    postcode: str
    country: str
    occupation: str
    company: str
    company_address: str
    company_contact: str
    emergency_contact: str
    emergency_phone: str
    billing_type: str
    remark: str
    medical_illness: str
    registered_on: datetime
    last_modified_by: str
    last_modified_on: datetime


@dataclass(frozen=True)
class PatientAllergy:
    substance: str
    recorded_by: str
    modified_on: datetime


@dataclass(frozen=True)
class PatientDocument:
    document_id: str
    title: str
    created_on: date
    effective_on: date
    created_by: str


@dataclass(frozen=True)
class PatientDeposit:
    deposit_id: str
    created_on: datetime
    amount: float
    transaction: str
    payment_code: str
    recorded_by: str
    remark: str


@dataclass(frozen=True)
class MedicReportSummary:
    report_id: int
    generated_on: datetime
    appointment_date: datetime
    created_by: str
    diagnosis: str
    treatment: str
    history: str
    examination: str
    finding: str
    advice: str
    next_action: str
    notes_preview: str


@dataclass(frozen=True)
class DentalCategory:
    id: int
    name: str
    status: str


@dataclass(frozen=True)
class DentalNotation:
    notation_id: int
    title: str
    stock_id: str
    stock_name: str
    procedure_desc: str
    price: float


@dataclass
class DentalChartItem:
    notation_id: int
    tooth_id: int
    tooth_plan: str = "E"
    remarks: str = ""
    unit_price: float = 0.0
    notation_status: int = 1
    bill_status: int = 0
    notation_title: str = ""
    stock_name: str = ""
    stock_id: str = ""


@dataclass(frozen=True)
class ReceiptDraftItem:
    stock_id: str
    description: str
    qty: int
    unit_price: float
    subtotal: float
    remark: str = ""


@dataclass(frozen=True)
class MedicReportDetail:
    report_id: int
    patient_id: str
    generated_on: datetime
    appointment_date: datetime
    created_by: str
    diagnosis: str
    treatment: str
    history: str
    examination: str
    finding: str
    advice: str
    next_action: str
    chart_items: List[DentalChartItem]


class ClinicDatabase:
    """Wrapper around the live MySQL database with query helpers."""

    def __init__(
        self,
        *,
        host: str = "localhost",
        port: int = 3306,
        user: str,
        password: str,
        database: str,
        charset: str = "utf8mb4",
    ) -> None:
        self._connection_kwargs = {
            "host": host,
            "port": int(port),
            "user": user,
            "password": password,
            "database": database,
        }
        self._conn = self._connect_with_charset(charset)

    def close(self) -> None:
        if getattr(self, "_conn", None):
            try:
                self._conn.close()
            finally:
                self._conn = None  # type: ignore[attr-defined]

    def _connect_with_charset(self, charset: str):
        try:
            return pymysql.connect(
                charset=charset,
                cursorclass=DictCursor,
                **self._connection_kwargs,
            )
        except pymysql_errors.OperationalError as exc:
            error_code = exc.args[0] if exc.args else None
            message = str(exc)
            if charset == "utf8mb4" and (error_code == 1115 or "Unknown character set" in message):
                try:
                    return pymysql.connect(
                        charset="utf8",
                        cursorclass=DictCursor,
                        **self._connection_kwargs,
                    )
                except pymysql_errors.MySQLError as fallback_exc:
                    raise RuntimeError(
                        "Unable to connect to MySQL using utf8mb4; fallback to utf8 failed"
                    ) from fallback_exc
            raise

    def _ensure_connection(self) -> None:
        if not getattr(self, "_conn", None):
            raise RuntimeError("Database connection not available")
        try:
            self._conn.ping(reconnect=True)
        except pymysql_errors.MySQLError as exc:
            raise RuntimeError(f"Lost MySQL connection: {exc}") from exc

    def _to_float(self, value) -> float:
        if value is None:
            return 0.0
        if isinstance(value, float):
            return value
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    def _to_datetime(self, value, default: Optional[datetime] = None) -> datetime:
        if isinstance(value, datetime):
            return value
        if isinstance(value, str):
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S"):
                try:
                    return datetime.strptime(value, fmt)
                except ValueError:
                    continue
        if isinstance(value, (int, float)):
            try:
                return datetime.fromtimestamp(value)
            except (OSError, OverflowError, ValueError):
                pass
        return default or datetime.min

    def _to_date(self, value, default: Optional[date] = None) -> date:
        if isinstance(value, date) and not isinstance(value, datetime):
            return value
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, str):
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d"):
                try:
                    return datetime.strptime(value, fmt).date()
                except ValueError:
                    continue
        return default or date(1900, 1, 1)



    def receipts_for_date(self, for_date: Optional[date], icpassport: Optional[str] = None) -> List[ReceiptSummary]:
        self._ensure_connection()
        sql = (
            "SELECT r.rcpt_id, r.issued, r.patient_id, r.total, r.subtotal, r.gst, r.payment AS payment_code, "
            "r.Remark AS remark, r.disc_total, r.rounding, r.consult_fees, r.done_by, r.department_type, "
            "r.settledby AS settled_by, r.tax_total, r.mr_id, "
            "p.icpassport AS p_ic, p.name AS p_name, p.receipt_name AS p_receipt_name, "
            "p.preferredname AS p_preferred_name, p.company AS p_company, "
            "p.phone_fixed AS p_phone_fixed, p.phone_mobile AS p_phone_mobile, p.Emailaddress AS p_email "
            "FROM receipts r JOIN patients p ON p.icpassport = r.patient_id "
            "WHERE r.removed = 0"
        )
        params: List[object] = []
        if for_date:
            start_dt = datetime.combine(for_date, datetime.min.time())
            end_dt = start_dt + timedelta(days=1)
            sql += " AND r.issued >= %s AND r.issued < %s"
            params.extend([start_dt, end_dt])
        if icpassport:
            sql += " AND p.icpassport = %s"
            params.append(icpassport.strip())
        sql += " ORDER BY r.issued DESC, p.name"

        with self._conn.cursor() as cursor:
            cursor.execute(sql, tuple(params))
            rows = cursor.fetchall()

        summaries: List[ReceiptSummary] = []
        for row in rows:
            issued = row["issued"]
            if isinstance(issued, str):
                issued = datetime.strptime(issued, "%Y-%m-%d %H:%M:%S")
            receipt = Receipt(
                rcpt_id=row["rcpt_id"],
                issued=issued,
                patient_id=row["patient_id"],
                total=self._to_float(row["total"]),
                subtotal=self._to_float(row["subtotal"]),
                gst=self._to_float(row["gst"]),
                payment_code=row.get("payment_code", ""),
                remark=row.get("remark", ""),
                discount=self._to_float(row.get("disc_total")),
                rounding=self._to_float(row.get("rounding")),
                consult_fees=self._to_float(row.get("consult_fees")),
                done_by=row.get("done_by", ""),
                department_type=row.get("department_type", ""),
                settled_by=row.get("settled_by", ""),
                tax_total=self._to_float(row.get("tax_total")),
                mr_id=int(row.get("mr_id", 0) or 0),
            )
            patient = Patient(
                icpassport=row.get("p_ic", ""),
                name=row.get("p_name", ""),
                receipt_name=row.get("p_receipt_name", ""),
                preferred_name=row.get("p_preferred_name", ""),
                company=row.get("p_company", ""),
                phone_fixed=row.get("p_phone_fixed", ""),
                phone_mobile=row.get("p_phone_mobile", ""),
                email=row.get("p_email", ""),
            )
            summaries.append(ReceiptSummary(receipt=receipt, patient=patient))
        return summaries

    def find_receipts(self, icpassport: str, for_date: date) -> List[Receipt]:
        return [summary.receipt for summary in self.receipts_for_date(for_date, icpassport)]

    def receipts_for_medic_report(self, mr_id: int) -> List[Receipt]:
        cleaned = int(mr_id or 0)
        if cleaned <= 0:
            return []
        self._ensure_connection()
        sql = (
            "SELECT rcpt_id, issued, patient_id, total, subtotal, gst, payment AS payment_code, "
            "Remark AS remark, disc_total, rounding, consult_fees, done_by, department_type, "
            "settledby AS settled_by, tax_total, mr_id "
            "FROM receipts "
            "WHERE removed = 0 AND mr_id = %s "
            "ORDER BY issued DESC"
        )
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            rows = cursor.fetchall()
        receipts: List[Receipt] = []
        for row in rows:
            issued = row.get("issued")
            issued_dt = self._to_datetime(issued) if not isinstance(issued, datetime) else issued
            receipts.append(
                Receipt(
                    rcpt_id=row.get("rcpt_id", ""),
                    issued=issued_dt,
                    patient_id=row.get("patient_id", ""),
                    total=self._to_float(row.get("total")),
                    subtotal=self._to_float(row.get("subtotal")),
                    gst=self._to_float(row.get("gst")),
                    payment_code=row.get("payment_code", ""),
                    remark=row.get("remark", ""),
                    discount=self._to_float(row.get("disc_total")),
                    rounding=self._to_float(row.get("rounding")),
                    consult_fees=self._to_float(row.get("consult_fees")),
                    done_by=row.get("done_by", ""),
                    department_type=row.get("department_type", ""),
                    settled_by=row.get("settled_by", ""),
                    tax_total=self._to_float(row.get("tax_total")),
                    mr_id=int(row.get("mr_id", 0) or 0),
                )
            )
        return receipts

    def get_receipt_items(self, rcpt_id: str) -> List[ReceiptItem]:
        self._ensure_connection()
        sql = (
            "SELECT ri.ID AS id, ri.rcpt_id, ri.item AS item_id, COALESCE(si.name, ri.item) AS name, ri.qty, "
            "ri.unitprice AS unit_price, ri.subtotal, ri.disc_amount, ri.username, ri.rcpt_remark AS remark "
            "FROM receipt_items ri "
            "LEFT JOIN stock_items si ON si.id = ri.item "
            "WHERE ri.rcpt_id = %s ORDER BY ri.ID"
        )
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (rcpt_id,))
            rows = cursor.fetchall()

        items: List[ReceiptItem] = []
        for row in rows:
            items.append(
                ReceiptItem(
                    id=int(row.get("id", 0)),
                    item_id=row.get("item_id", ""),
                    name=row.get("name", ""),
                    qty=int(row.get("qty", 0) or 0),
                    unit_price=self._to_float(row.get("unit_price")),
                    subtotal=self._to_float(row.get("subtotal")),
                    discount=self._to_float(row.get("disc_amount")),
                    username=row.get("username", ""),
                    remark=row.get("remark", ""),
                )
            )
        return items






    def partial_payments_for_receipts(self, receipt_ids: Sequence[str]) -> Dict[str, List[PartialPayment]]:
        if not receipt_ids:
            return {}

        self._ensure_connection()
        placeholders = ", ".join(["%s"] * len(receipt_ids))
        sql = (
            "SELECT pp.payment_id, pp.rcpt_id AS rcpt_id, pp.date, pp.amount, pp.pay_code, "
            "COALESCE(pm.description, pp.pay_code) AS method_desc, pp.username, pp.settledby, pp.remark "
            "FROM partial_payment pp "
            "LEFT JOIN payment_method pm ON pm.paycode = pp.pay_code "
            "WHERE pp.removed = 0 AND pp.rcpt_id IN ("
            + placeholders
            + ") ORDER BY pp.rcpt_id, pp.date, pp.payment_id"
        )
        with self._conn.cursor() as cursor:
            cursor.execute(sql, tuple(receipt_ids))
            rows = cursor.fetchall()

        payments: Dict[str, List[PartialPayment]] = {rid: [] for rid in receipt_ids}
        for row in rows:
            date_value = row.get("date")
            if isinstance(date_value, str):
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
                    try:
                        date_value = datetime.strptime(date_value, fmt)
                        break
                    except ValueError:
                        continue
                else:
                    date_value = datetime.min
            elif not isinstance(date_value, datetime):
                date_value = datetime.min

            amount = self._to_float(row.get("amount"))
            receipt_id = row.get("rcpt_id", "")
            payment = PartialPayment(
                payment_id=int(row.get("payment_id", 0) or 0),
                receipt_id=receipt_id,
                date=date_value,
                amount=amount,
                pay_code=row.get("pay_code", ""),
                method=row.get("method_desc") or row.get("pay_code", ""),
                username=row.get("username", ""),
                settled_by=row.get("settledby", ""),
                remark=row.get("remark", ""),
            )
            payments.setdefault(receipt_id, []).append(payment)
        return payments


    def partial_payments_for_recripts(self, receipt_ids: Sequence[str]) -> Dict[str, List[PartialPayment]]:
        return self.partial_payments_for_receipts(receipt_ids)

    def update_partial_payment_amount(self, payment_id: int, *, amount: float) -> None:
        self._ensure_connection()
        cursor = self._conn.cursor()
        try:
            self._conn.begin()
            cursor.execute(
                "UPDATE partial_payment SET amount = %s WHERE payment_id = %s AND removed = 0",
                (float(amount), int(payment_id)),
            )
            self._conn.commit()
        except Exception:
            self._conn.rollback()
            raise
        finally:
            cursor.close()

    def get_payment_description(self, paycode: str) -> str:
        self._ensure_connection()
        sql = "SELECT description FROM payment_method WHERE paycode = %s"
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (paycode.strip(),))
            row = cursor.fetchone()
        if row and row.get("description"):
            return row["description"]
        return paycode


    def all_payment_codes(self) -> Dict[str, str]:
        self._ensure_connection()
        sql = "SELECT paycode, description FROM payment_method ORDER BY description"
        with self._conn.cursor() as cursor:
            cursor.execute(sql)
            rows = cursor.fetchall()
        return {str(row["paycode"]): str(row["description"]) for row in rows}


    def appointments_for_date(self, for_date: date) -> List[AppointmentDetail]:
        self._ensure_connection()
        target_date = for_date or date.today()
        start_dt = datetime.combine(target_date, datetime.min.time())
        end_dt = start_dt + timedelta(days=1)
        sql = """
            SELECT
                a.id,
                a.apt_date,
                a.patient_id,
                COALESCE(NULLIF(p.receipt_name, ''), NULLIF(p.preferredname, ''), NULLIF(p.name, ''), a.patient_id) AS patient_name,
                a.reason,
                a.resource,
                a.type AS appointment_type,
                a.location,
                COALESCE(NULLIF(a.queue_number, ''), '-') AS queue_number,
                a.username AS created_by,
                COALESCE(s.status_desc, 'Scheduled') AS status_desc,
                COALESCE(NULLIF(s.slot_view_color, ''), '#FFFFFF') AS status_color,
                COALESCE(last_status.status_id, 0) AS status_id
            FROM appointments a
            LEFT JOIN patients p ON p.icpassport = a.patient_id
            LEFT JOIN (
                SELECT t.apt_date, t.patient_id, t.status_id
                FROM appointment_tracking t
                JOIN (
                    SELECT apt_date, patient_id, MAX(created_time) AS max_created_time
                    FROM appointment_tracking
                    GROUP BY apt_date, patient_id
                ) latest
                  ON latest.apt_date = t.apt_date
                 AND latest.patient_id = t.patient_id
                 AND latest.max_created_time = t.created_time
            ) last_status
              ON last_status.apt_date = a.apt_date AND last_status.patient_id = a.patient_id
            LEFT JOIN appointment_status s ON s.id = last_status.status_id
            WHERE a.removed = 0
              AND a.apt_date >= %s
              AND a.apt_date < %s
            ORDER BY a.apt_date ASC, patient_name ASC
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (start_dt, end_dt))
            rows = cursor.fetchall()

        details: List[AppointmentDetail] = []
        for row in rows:
            when = row.get("apt_date")
            if isinstance(when, str):
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                    try:
                        when = datetime.strptime(when, fmt)
                        break
                    except ValueError:
                        continue
                else:
                    when = datetime.min
            elif not isinstance(when, datetime):
                when = datetime.min

            details.append(
                AppointmentDetail(
                    id=int(row.get("id", 0) or 0),
                    scheduled=when,
                    patient_id=row.get("patient_id", "") or "",
                    patient_name=row.get("patient_name", "") or "",
                    reason=row.get("reason", "") or "",
                    resource=row.get("resource", "") or "",
                    appointment_type=row.get("appointment_type", "") or "",
                    location=row.get("location", "") or "",
                    queue_number=row.get("queue_number", "") or "",
                    created_by=row.get("created_by", "") or "",
                    status=row.get("status_desc", "") or "",
                    status_color=row.get("status_color", "#FFFFFF") or "#FFFFFF",
                    status_id=int(row.get("status_id", 0) or 0),
                )
            )
        return details

    def appointments_for_status(self, for_date: date, status_ids: Sequence[int]) -> List[AppointmentDetail]:
        if not status_ids:
            return []
        self._ensure_connection()
        target_date = for_date or date.today()
        start_dt = datetime.combine(target_date, datetime.min.time())
        end_dt = start_dt + timedelta(days=1)
        placeholders = ", ".join(["%s"] * len(status_ids))
        sql = f"""
            SELECT
                a.id,
                a.apt_date,
                a.patient_id,
                COALESCE(NULLIF(p.receipt_name, ''), NULLIF(p.preferredname, ''), NULLIF(p.name, ''), a.patient_id) AS patient_name,
                a.reason,
                a.resource,
                a.type AS appointment_type,
                a.location,
                COALESCE(NULLIF(a.queue_number, ''), '-') AS queue_number,
                a.username AS created_by,
                COALESCE(s.status_desc, 'Scheduled') AS status_desc,
                COALESCE(NULLIF(s.slot_view_color, ''), '#FFFFFF') AS status_color,
                COALESCE(last_status.status_id, 0) AS status_id
            FROM appointments a
            LEFT JOIN patients p ON p.icpassport = a.patient_id
            LEFT JOIN (
                SELECT t.apt_date, t.patient_id, t.status_id
                FROM appointment_tracking t
                JOIN (
                    SELECT apt_date, patient_id, MAX(created_time) AS max_created_time
                    FROM appointment_tracking
                    GROUP BY apt_date, patient_id
                ) latest
                  ON latest.apt_date = t.apt_date
                 AND latest.patient_id = t.patient_id
                 AND latest.max_created_time = t.created_time
            ) last_status
              ON last_status.apt_date = a.apt_date AND last_status.patient_id = a.patient_id
            LEFT JOIN appointment_status s ON s.id = last_status.status_id
            WHERE a.removed = 0
              AND a.apt_date >= %s
              AND a.apt_date < %s
              AND COALESCE(last_status.status_id, 0) IN ({placeholders})
            ORDER BY a.apt_date ASC, patient_name ASC
        """
        params: List[object] = [start_dt, end_dt]
        params.extend(int(s) for s in status_ids)
        with self._conn.cursor() as cursor:
            cursor.execute(sql, tuple(params))
            rows = cursor.fetchall()

        details: List[AppointmentDetail] = []
        for row in rows:
            when = row.get("apt_date")
            if isinstance(when, str):
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
                    try:
                        when = datetime.strptime(when, fmt)
                        break
                    except ValueError:
                        continue
                else:
                    when = datetime.min
            elif not isinstance(when, datetime):
                when = datetime.min

            details.append(
                AppointmentDetail(
                    id=int(row.get("id", 0) or 0),
                    scheduled=when,
                    patient_id=row.get("patient_id", "") or "",
                    patient_name=row.get("patient_name", "") or "",
                    reason=row.get("reason", "") or "",
                    resource=row.get("resource", "") or "",
                    appointment_type=row.get("appointment_type", "") or "",
                    location=row.get("location", "") or "",
                    queue_number=row.get("queue_number", "") or "",
                    created_by=row.get("created_by", "") or "",
                    status=row.get("status_desc", "") or "",
                    status_color=row.get("status_color", "#FFFFFF") or "#FFFFFF",
                    status_id=int(row.get("status_id", 0) or 0),
                )
            )
        return details

    def get_patient_profile(self, patient_id: str) -> Optional[PatientProfile]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return None
        self._ensure_connection()
        sql = """
            SELECT
                register_date,
                icpassport,
                name,
                preferredname,
                receipt_name,
                sex,
                DOB,
                phone_mobile,
                phone_fixed,
                Emailaddress,
                address,
                city,
                state,
                postcode,
                country,
                occupation,
                company,
                companyaddress,
                companycontact,
                EmergencyContact,
                EmergencyPhoneNo,
                BillingType,
                remark,
                medical_illness,
                username,
                modified_date
            FROM patients
            WHERE icpassport = %s AND removed = 0
            LIMIT 1
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            row = cursor.fetchone()
        if not row:
            return None
        dob = self._to_date(row.get("DOB"))
        register_date = self._to_datetime(row.get("register_date"))
        modified_on = self._to_datetime(row.get("modified_date"))
        return PatientProfile(
            patient_id=row.get("icpassport", "") or "",
            name=row.get("name", "") or "",
            preferred_name=row.get("preferredname", "") or "",
            receipt_name=row.get("receipt_name", "") or "",
            sex=row.get("sex", "") or "",
            date_of_birth=dob,
            phone_mobile=row.get("phone_mobile", "") or "",
            phone_fixed=row.get("phone_fixed", "") or "",
            email=row.get("Emailaddress", "") or "",
            address=row.get("address", "") or "",
            city=row.get("city", "") or "",
            state=row.get("state", "") or "",
            postcode=row.get("postcode", "") or "",
            country=row.get("country", "") or "",
            occupation=row.get("occupation", "") or "",
            company=row.get("company", "") or "",
            company_address=row.get("companyaddress", "") or "",
            company_contact=row.get("companycontact", "") or "",
            emergency_contact=row.get("EmergencyContact", "") or "",
            emergency_phone=row.get("EmergencyPhoneNo", "") or "",
            billing_type=row.get("BillingType", "") or "",
            remark=row.get("remark", "") or "",
            medical_illness=row.get("medical_illness", "") or "",
            registered_on=register_date,
            last_modified_by=row.get("username", "") or "",
            last_modified_on=modified_on,
        )

    def allergies_for_patient(self, patient_id: str) -> List[PatientAllergy]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return []
        self._ensure_connection()
        sql = """
            SELECT drug_name, username, modified_date
            FROM allergy
            WHERE patient_id = %s
            ORDER BY modified_date DESC, drug_name
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            rows = cursor.fetchall()
        allergies: List[PatientAllergy] = []
        for row in rows:
            allergies.append(
                PatientAllergy(
                    substance=row.get("drug_name", "") or "",
                    recorded_by=row.get("username", "") or "",
                    modified_on=self._to_datetime(row.get("modified_date")),
                )
            )
        return allergies

    def documents_for_patient(self, patient_id: str, limit: int = 50) -> List[PatientDocument]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return []
        limit = max(1, min(limit, 200))
        self._ensure_connection()
        sql = """
            SELECT id, title, created, effective, username
            FROM patientdocs
            WHERE patientID = %s
            ORDER BY modified_date DESC
            LIMIT %s
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned, limit))
            rows = cursor.fetchall()
        documents: List[PatientDocument] = []
        for row in rows:
            documents.append(
                PatientDocument(
                    document_id=row.get("id", "") or "",
                    title=row.get("title", "") or "",
                    created_on=self._to_date(row.get("created")),
                    effective_on=self._to_date(row.get("effective")),
                    created_by=row.get("username", "") or "",
                )
            )
        return documents

    def deposits_for_patient(self, patient_id: str) -> Tuple[List[PatientDeposit], float]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return ([], 0.0)
        self._ensure_connection()
        sql = """
            SELECT deposit_id, created_date, amount, transaction, paymentmethod, username, Remark
            FROM deposit
            WHERE patient_id = %s AND removed = 0
            ORDER BY created_date DESC
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            rows = cursor.fetchall()
        deposits: List[PatientDeposit] = []
        total = 0.0
        for row in rows:
            amount = self._to_float(row.get("amount"))
            total += amount
            deposits.append(
                PatientDeposit(
                    deposit_id=row.get("deposit_id", "") or "",
                    created_on=self._to_datetime(row.get("created_date")),
                    amount=amount,
                    transaction=row.get("transaction", "") or "",
                    payment_code=row.get("paymentmethod", "") or "",
                    recorded_by=row.get("username", "") or "",
                    remark=row.get("Remark", "") or "",
                )
            )
        return deposits, total

    def medic_reports_for_patient(self, patient_id: str, limit: int = 50) -> List[MedicReportSummary]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return []
        limit = max(1, min(limit, 200))
        self._ensure_connection()
        sql = """
            SELECT
                id,
                generated_date,
                apt_date,
                username,
                diagnosis,
                treatment,
                COALESCE(history, '') AS history,
                COALESCE(examination, '') AS examination,
                COALESCE(finding, '') AS finding,
                COALESCE(advice, '') AS advice,
                COALESCE(nextAction, '') AS next_action
            FROM medic_report
            WHERE patient_id = %s
            ORDER BY generated_date DESC
            LIMIT %s
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned, limit))
            rows = cursor.fetchall()

        summaries: List[MedicReportSummary] = []
        for row in rows:
            history = (row.get("history") or "").strip()
            examination = (row.get("examination") or "").strip()
            finding = (row.get("finding") or "").strip()
            advice = (row.get("advice") or "").strip()
            next_action = (row.get("next_action") or "").strip()
            summaries.append(
                MedicReportSummary(
                    report_id=int(row.get("id", 0) or 0),
                    generated_on=self._to_datetime(row.get("generated_date")),
                    appointment_date=self._to_datetime(row.get("apt_date")),
                    created_by=row.get("username", "") or "",
                    diagnosis=(row.get("diagnosis") or "").strip(),
                    treatment=(row.get("treatment") or "").strip(),
                    history=history,
                    examination=examination,
                    finding=finding,
                    advice=advice,
                    next_action=next_action,
                    notes_preview="",
                )
            )
        return summaries

    def dental_categories(self) -> List[DentalCategory]:
        self._ensure_connection()
        sql = """
            SELECT categoryid, categorydesc, categorystatus
            FROM dentalnotation_category
            ORDER BY categorydesc
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql)
            rows = cursor.fetchall()
        categories: List[DentalCategory] = []
        for row in rows:
            categories.append(
                DentalCategory(
                    id=int(row.get("categoryid", 0) or 0),
                    name=row.get("categorydesc", "") or "",
                    status=row.get("categorystatus", "") or "",
                )
            )
        return categories

    def dental_notations(self, category_id: int) -> List[DentalNotation]:
        self._ensure_connection()
        sql = """
            SELECT
                cfg.notationid,
                cfg.notationtitle,
                cfg.stock_id,
                COALESCE(si.name, '') AS stock_name,
                COALESCE(si.`procedure`, '') AS procedure_desc,
                COALESCE(si.selling_price, 0.0) AS selling_price
            FROM dentalnotation_config cfg
            LEFT JOIN stock_items si ON si.id = cfg.stock_id
            WHERE cfg.categoryid = %s
            ORDER BY cfg.notationtitle
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (int(category_id),))
            rows = cursor.fetchall()
        notations: List[DentalNotation] = []
        for row in rows:
            price = row.get("selling_price")
            notations.append(
                DentalNotation(
                    notation_id=int(row.get("notationid", 0) or 0),
                    title=row.get("notationtitle", "") or "",
                    stock_id=row.get("stock_id", "") or "",
                    stock_name=row.get("stock_name", "") or "",
                    procedure_desc=row.get("procedure_desc", "") or "",
                    price=float(price or 0.0),
                )
            )
        return notations

    def create_medical_record(
        self,
        *,
        patient_id: str,
        generated_on: datetime,
        appointment_on: Optional[datetime],
        username: str,
        history: str = "",
        diagnosis: str = "",
        treatment: str = "",
        examination: str = "",
        finding: str = "",
        advice: str = "",
        next_action: str = "",
        chart_items: Optional[Sequence[DentalChartItem]] = None,
    ) -> int:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            raise ValueError("patient_id is required")
        self._ensure_connection()
        generated = generated_on or datetime.now()
        appointment_dt = appointment_on or generated
        chart_items = list(chart_items or [])
        insert_sql = """
            INSERT INTO medic_report (
                generated_date,
                patient_id,
                history,
                diagnosis,
                treatment,
                examination,
                finding,
                advice,
                nextAction,
                username,
                modified_date,
                apt_date
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        chart_sql = """
            INSERT INTO medic_report_dentalchart (
                parentid,
                toothid,
                toothplan,
                notationid,
                notationstatus,
                remarks,
                unitprice,
                billstatus,
                username,
                modified_date
            ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor = self._conn.cursor()
        try:
            cursor.execute(
                insert_sql,
                (
                    generated,
                    cleaned,
                    history,
                    diagnosis,
                    treatment,
                    examination,
                    finding,
                    advice,
                    next_action,
                    username,
                    generated,
                    appointment_dt,
                ),
            )
            report_id = cursor.lastrowid
            if chart_items:
                payload = []
                for item in chart_items:
                    payload.append(
                        (
                            report_id,
                            int(item.tooth_id),
                            (item.tooth_plan or "E")[:1],
                            int(item.notation_id),
                            int(item.notation_status),
                            item.remarks or "",
                            float(item.unit_price or 0.0),
                            int(item.bill_status),
                            username,
                            generated,
                        )
                    )
                cursor.executemany(chart_sql, payload)
            self._conn.commit()
            return int(report_id)
        except Exception:
            self._conn.rollback()
            raise
        finally:
            cursor.close()

    def _allocate_receipt_id(self, cursor, issued: datetime) -> Tuple[str, int]:
        cursor.execute(
            """
            SELECT COALESCE(MAX(CAST(SUBSTRING(SUBSTRING_INDEX(rcpt_id, '/', 1), 2) AS UNSIGNED)), 0) AS seq
            FROM receipts
            FOR UPDATE
            """
        )
        row = cursor.fetchone() or {}
        last_seq = int(row.get("seq", 0) or 0)
        next_seq = last_seq + 1
        rcpt_id = f"A{next_seq:06d}/{issued.year}"
        return rcpt_id, next_seq

    def create_receipt(
        self,
        *,
        patient_id: str,
        issued: datetime,
        username: str,
        payment_code: str,
        items: Sequence[ReceiptDraftItem],
        subtotal: float,
        discount: float,
        rounding: float,
        consult_fees: float = 0.0,
        remark: str = "",
        mr_id: int = 0,
        department: str = "D1",
    ) -> str:
        if not items:
            raise ValueError("At least one receipt item is required")
        cleaned_patient = (patient_id or "").strip()
        if not cleaned_patient:
            raise ValueError("patient_id is required")
        self._ensure_connection()
        subtotal_value = float(subtotal or 0.0)
        consult_value = float(consult_fees or 0.0)
        discount_value = float(discount or 0.0)
        rounding_value = float(rounding or 0.0)
        total = subtotal_value + consult_value - discount_value + rounding_value
        cursor = self._conn.cursor()
        try:
            self._conn.begin()
            rcpt_id, seq = self._allocate_receipt_id(cursor, issued)
            receipt_sql = """
                INSERT INTO receipts (
                    rcpt_id,
                    issued,
                    patient_id,
                    consult_fees,
                    total,
                    payment,
                    Billed,
                    Remark,
                    removed,
                    subtotal,
                    gst,
                    itemize,
                    username,
                    modified_date,
                    settledby,
                    tax_total,
                    disc_total,
                    done_by,
                    rounding,
                    mr_id,
                    department_type,
                    printed_name
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 0, %s, %s, 1, %s, %s, %s, %s, %s, %s, %s, %s, %s, NULL)
            """
            cursor.execute(
                receipt_sql,
                (
                    rcpt_id,
                    issued,
                    cleaned_patient,
                    consult_value,
                    total,
                    payment_code,
                    "00",
                    remark or "",
                    subtotal_value,
                    0.0,
                    username,
                    issued,
                    username,
                    0.0,
                    discount_value,
                    username,
                    rounding_value,
                    int(mr_id or 0),
                    department,
                ),
            )

            item_sql = """
                INSERT INTO receipt_items (
                    rcpt_id,
                    item,
                    qty,
                    unitprice,
                    subtotal,
                    rcpt_remark,
                    username,
                    modified_date,
                    doseqty,
                    timesqty,
                    bodypart,
                    disc_amount
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 0, 0, '', 0)
            """
            modified = datetime.now()
            payload = []
            for item in items:
                payload.append(
                    (
                        rcpt_id,
                        item.stock_id,
                        int(item.qty),
                        float(item.unit_price),
                        float(item.subtotal),
                        item.remark or item.description or "",
                        username,
                        modified,
                    )
                )
            cursor.executemany(item_sql, payload)

            cursor.execute(
                "INSERT INTO log_receipt (rcpt_id, seq, company_code, total) VALUES (%s, %s, %s, %s)",
                (rcpt_id, seq, department, total),
            )

            self._conn.commit()
            return rcpt_id
        except Exception:
            self._conn.rollback()
            raise
        finally:
            cursor.close()

    def replace_receipt(
        self,
        rcpt_id: str,
        *,
        issued: datetime,
        patient_id: str,
        items: Sequence[ReceiptDraftItem],
        subtotal: float,
        discount: float,
        rounding: float,
        consult_fees: float,
        remark: str,
        payment_code: str,
        username: str,
        department: str,
        mr_id: int,
    ) -> None:
        if not items:
            raise ValueError("At least one receipt item is required")
        cleaned_rcpt = (rcpt_id or "").strip()
        if not cleaned_rcpt:
            raise ValueError("rcpt_id is required")
        cleaned_patient = (patient_id or "").strip()
        if not cleaned_patient:
            raise ValueError("patient_id is required")
        self._ensure_connection()
        subtotal_value = float(subtotal or 0.0)
        consult_value = float(consult_fees or 0.0)
        discount_value = float(discount or 0.0)
        rounding_value = float(rounding or 0.0)
        total = subtotal_value + consult_value - discount_value + rounding_value
        cursor = self._conn.cursor()
        try:
            self._conn.begin()
            cursor.execute(
                """
                UPDATE receipts
                SET issued = %s,
                    patient_id = %s,
                    consult_fees = %s,
                    total = %s,
                    payment = %s,
                    Remark = %s,
                    subtotal = %s,
                    gst = 0.0,
                    itemize = 1,
                    username = %s,
                    modified_date = %s,
                    settledby = %s,
                    tax_total = 0.0,
                    disc_total = %s,
                    done_by = %s,
                    rounding = %s,
                    mr_id = %s,
                    department_type = %s
                WHERE rcpt_id = %s
                """,
                (
                    issued,
                    cleaned_patient,
                    consult_value,
                    total,
                    payment_code,
                    remark or "",
                    subtotal_value,
                    username,
                    issued,
                    username,
                    discount_value,
                    username,
                    rounding_value,
                    int(mr_id or 0),
                    department,
                    cleaned_rcpt,
                ),
            )
            cursor.execute("DELETE FROM receipt_items WHERE rcpt_id = %s", (cleaned_rcpt,))
            insert_sql = """
                INSERT INTO receipt_items (
                    rcpt_id,
                    item,
                    qty,
                    unitprice,
                    subtotal,
                    rcpt_remark,
                    username,
                    modified_date,
                    doseqty,
                    timesqty,
                    bodypart,
                    disc_amount
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, 0, 0, '', 0)
            """
            modified = datetime.now()
            payload = []
            for item in items:
                payload.append(
                    (
                        cleaned_rcpt,
                        item.stock_id,
                        int(item.qty),
                        float(item.unit_price),
                        float(item.subtotal),
                        item.remark or item.description or "",
                        username,
                        modified,
                    )
                )
            cursor.executemany(insert_sql, payload)
            try:
                cursor.execute("UPDATE log_receipt SET total = %s WHERE rcpt_id = %s", (total, cleaned_rcpt))
                if cursor.rowcount == 0:
                    cursor.execute(
                        "INSERT INTO log_receipt (rcpt_id, seq, company_code, total) VALUES (%s, %s, %s, %s)",
                        (cleaned_rcpt, 0, department, total),
                    )
            except Exception:
                pass
            self._conn.commit()
        except Exception:
            self._conn.rollback()
            raise
        finally:
            cursor.close()

    def record_appointment_status(
        self,
        patient_id: str,
        appointment_time: datetime,
        status_id: int,
        username: str = "",
    ) -> None:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return
        self._ensure_connection()
        sql = """
            INSERT INTO appointment_tracking (
                apt_date,
                patient_id,
                created_time,
                status_id,
                duration,
                username,
                modified_date
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s)
        """
        now = datetime.now()
        with self._conn.cursor() as cursor:
            cursor.execute(
                sql,
                (
                    appointment_time,
                    cleaned,
                    now,
                    int(status_id),
                    0,
                    username or None,
                    now,
                ),
            )
        self._conn.commit()

    def payment_codes_list(self) -> List[Tuple[str, str]]:
        methods = self.all_payment_codes()
        return sorted(methods.items(), key=lambda x: x[1])

    def payment_methods_list(self) -> List[Tuple[str, str]]:
        """
        Alias for payment codes, retained for UI compatibility.
        Returns (code, description) pairs sorted by description.
        """
        return self.payment_codes_list()

    def stock_items_search(self, query: str, limit: int = 50) -> List[Tuple[str, str, float]]:
        self._ensure_connection()
        q = f"%{(query or '').strip()}%"
        sql = """
            SELECT id, name, COALESCE(selling_price, 0.0) AS price
            FROM stock_items
            WHERE removed = 0 AND (name LIKE %s OR id LIKE %s OR procedure LIKE %s)
            ORDER BY name ASC
            LIMIT %s
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (q, q, q, max(1, min(limit, 200))))
            rows = cursor.fetchall()
        return [(row["id"], row["name"], float(row.get("price") or 0.0)) for row in rows]

    def stock_item_details(self, stock_id: str) -> Optional[Tuple[str, str, float]]:
        cleaned = (stock_id or "").strip()
        if not cleaned:
            return None
        self._ensure_connection()
        sql = "SELECT id, name, COALESCE(selling_price, 0.0) AS price FROM stock_items WHERE id = %s LIMIT 1"
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            row = cursor.fetchone()
        if not row:
            return None
        return row["id"], row["name"], float(row.get("price") or 0.0)

    def stock_categories(self) -> List[str]:
        self._ensure_connection()
        sql = "SELECT DISTINCT category FROM stock_items WHERE category IS NOT NULL AND TRIM(category) != '' ORDER BY category"
        with self._conn.cursor() as cursor:
            cursor.execute(sql)
            rows = cursor.fetchall()
        return [row["category"] for row in rows if row.get("category")]

    def stock_items_by_category(self, category: str) -> List[Tuple[str, str, float]]:
        self._ensure_connection()
        sql = """
            SELECT id, name, COALESCE(selling_price, 0.0) AS price
            FROM stock_items
            WHERE category = %s AND removed = 0
            ORDER BY name
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (category,))
            rows = cursor.fetchall()
        return [(row["id"], row["name"], float(row.get("price") or 0.0)) for row in rows]

    def _chart_items_for_report(self, report_id: int) -> List[DentalChartItem]:
        self._ensure_connection()
        sql = """
            SELECT
                dc.notationid,
                dc.toothid,
                dc.toothplan,
                dc.remarks,
                dc.unitprice,
                dc.notationstatus,
                dc.billstatus,
                cfg.notationtitle,
                cfg.stock_id,
                COALESCE(si.name, '') AS stock_name
            FROM medic_report_dentalchart dc
            LEFT JOIN dentalnotation_config cfg ON cfg.notationid = dc.notationid
            LEFT JOIN stock_items si ON si.id = cfg.stock_id
            WHERE dc.parentid = %s
            ORDER BY dc.id
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (int(report_id),))
            rows = cursor.fetchall()
        items: List[DentalChartItem] = []
        for row in rows:
            items.append(
                DentalChartItem(
                    notation_id=int(row.get("notationid", 0) or 0),
                    tooth_id=int(row.get("toothid", 0) or 0),
                    tooth_plan=(row.get("toothplan") or "E")[:1].upper(),
                    remarks=row.get("remarks", "") or "",
                    unit_price=self._to_float(row.get("unitprice")),
                    notation_status=int(row.get("notationstatus", 1) or 1),
                    bill_status=int(row.get("billstatus", 0) or 0),
                    notation_title=row.get("notationtitle", "") or "",
                    stock_name=row.get("stock_name", "") or "",
                    stock_id=row.get("stock_id", "") or "",
                )
            )
        return items

    def recent_receipts_for_patient(self, patient_id: str, limit: int = 20) -> List[ReceiptSummary]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return []
        self._ensure_connection()
        sql = (
            "SELECT r.rcpt_id, r.issued, r.patient_id, r.total, r.subtotal, r.gst, r.payment AS payment_code, "
            "r.Remark AS remark, r.disc_total, r.rounding, r.consult_fees, r.done_by, r.department_type, "
            "r.settledby AS settled_by, r.tax_total, "
            "p.icpassport AS p_ic, p.name AS p_name, p.receipt_name AS p_receipt_name, "
            "p.preferredname AS p_preferred_name, p.company AS p_company, "
            "p.phone_fixed AS p_phone_fixed, p.phone_mobile AS p_phone_mobile, p.Emailaddress AS p_email "
            "FROM receipts r JOIN patients p ON p.icpassport = r.patient_id "
            "WHERE r.removed = 0 AND r.patient_id = %s "
            "ORDER BY r.issued DESC LIMIT %s"
        )
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned, int(max(1, limit))))
            rows = cursor.fetchall()

        summaries: List[ReceiptSummary] = []
        for row in rows:
            issued = row["issued"]
            if isinstance(issued, str):
                issued = datetime.strptime(issued, "%Y-%m-%d %H:%M:%S")
            receipt = Receipt(
                rcpt_id=row["rcpt_id"],
                issued=issued,
                patient_id=row["patient_id"],
                total=self._to_float(row["total"]),
                subtotal=self._to_float(row["subtotal"]),
                gst=self._to_float(row["gst"]),
                payment_code=row.get("payment_code", ""),
                remark=row.get("remark", ""),
                discount=self._to_float(row.get("disc_total")),
                rounding=self._to_float(row.get("rounding")),
                consult_fees=self._to_float(row.get("consult_fees")),
                done_by=row.get("done_by", ""),
                department_type=row.get("department_type", ""),
                settled_by=row.get("settled_by", ""),
                tax_total=self._to_float(row.get("tax_total")),
            )
            patient = Patient(
                icpassport=row.get("p_ic", ""),
                name=row.get("p_name", ""),
                receipt_name=row.get("p_receipt_name", ""),
                preferred_name=row.get("p_preferred_name", ""),
                company=row.get("p_company", ""),
                phone_fixed=row.get("p_phone_fixed", ""),
                phone_mobile=row.get("p_phone_mobile", ""),
                email=row.get("p_email", ""),
            )
            summaries.append(ReceiptSummary(receipt=receipt, patient=patient))
        return summaries

    def receipt_items_for_patient_date(self, patient_id: str, for_date: date) -> List[Tuple[Receipt, List[ReceiptItem]]]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return []
        target_date = for_date or date.today()
        summaries = self.receipts_for_date(target_date, cleaned)
        result: List[Tuple[Receipt, List[ReceiptItem]]] = []
        for summary in summaries:
            items = self.get_receipt_items(summary.receipt.rcpt_id)
            result.append((summary.receipt, items))
        return result

    def notation_for_stock(self, stock_id: str) -> Optional[DentalNotation]:
        cleaned = (stock_id or "").strip()
        if not cleaned:
            return None
        self._ensure_connection()
        sql = """
            SELECT
                cfg.notationid,
                cfg.notationtitle,
                cfg.stock_id,
                COALESCE(si.name, '') AS stock_name,
                COALESCE(si.`procedure`, '') AS procedure_desc,
                COALESCE(si.selling_price, 0.0) AS selling_price
            FROM dentalnotation_config cfg
            LEFT JOIN stock_items si ON si.id = cfg.stock_id
            WHERE cfg.stock_id = %s
            ORDER BY cfg.notationid
            LIMIT 1
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            row = cursor.fetchone()
        if not row:
            return None
        return DentalNotation(
            notation_id=int(row.get("notationid", 0) or 0),
            title=row.get("notationtitle", "") or "",
            stock_id=row.get("stock_id", "") or "",
            stock_name=row.get("stock_name", "") or "",
            procedure_desc=row.get("procedure_desc", "") or "",
            price=float(row.get("selling_price") or 0.0),
        )


    def notation_by_id(self, notation_id: int) -> Optional[DentalNotation]:
        self._ensure_connection()
        sql = """
            SELECT
                cfg.notationid,
                cfg.notationtitle,
                cfg.stock_id,
                COALESCE(si.name, '') AS stock_name,
                COALESCE(si.`procedure`, '') AS procedure_desc,
                COALESCE(si.selling_price, 0.0) AS selling_price
            FROM dentalnotation_config cfg
            LEFT JOIN stock_items si ON si.id = cfg.stock_id
            WHERE cfg.notationid = %s
            LIMIT 1
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (int(notation_id),))
            row = cursor.fetchone()
        if not row:
            return None
        return DentalNotation(
            notation_id=int(row.get('notationid', 0) or 0),
            title=row.get('notationtitle', '') or '',
            stock_id=row.get('stock_id', '') or '',
            stock_name=row.get('stock_name', '') or '',
            procedure_desc=row.get('procedure_desc', '') or '',
            price=float(row.get('selling_price') or 0.0),
        )

    def medic_report_for_appointment(
        self, patient_id: str, appointment_dt: datetime
    ) -> Optional[MedicReportDetail]:
        cleaned = (patient_id or "").strip()
        if not cleaned:
            return None
        self._ensure_connection()
        sql = """
            SELECT
                id,
                generated_date,
                apt_date,
                username,
                diagnosis,
                treatment,
                history,
                examination,
                finding,
                advice,
                nextAction
            FROM medic_report
            WHERE patient_id = %s
              AND apt_date BETWEEN (%s - INTERVAL 6 HOUR) AND (%s + INTERVAL 6 HOUR)
            ORDER BY ABS(TIMESTAMPDIFF(MINUTE, apt_date, %s)) ASC, generated_date DESC
            LIMIT 1
        """
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned, appointment_dt, appointment_dt, appointment_dt))
            row = cursor.fetchone()
        if not row:
            return None
        report_id = int(row.get("id", 0) or 0)
        items = self._chart_items_for_report(report_id)
        return MedicReportDetail(
            report_id=report_id,
            patient_id=cleaned,
            generated_on=self._to_datetime(row.get("generated_date")),
            appointment_date=self._to_datetime(row.get("apt_date")),
            created_by=row.get("username", "") or "",
            diagnosis=(row.get("diagnosis") or ""),
            treatment=(row.get("treatment") or ""),
            history=(row.get("history") or ""),
            examination=(row.get("examination") or ""),
            finding=(row.get("finding") or ""),
            advice=(row.get("advice") or ""),
            next_action=(row.get("nextAction") or ""),
            chart_items=items,
        )

    def update_medical_record(
        self,
        report_id: int,
        *,
        generated_on: datetime,
        appointment_on: datetime,
        username: str,
        history: str = "",
        diagnosis: str = "",
        treatment: str = "",
        examination: str = "",
        finding: str = "",
        advice: str = "",
        next_action: str = "",
        chart_items: Optional[Sequence[DentalChartItem]] = None,
    ) -> None:
        self._ensure_connection()
        chart_items = list(chart_items or [])
        update_sql = """
            UPDATE medic_report
            SET generated_date = %s,
                apt_date = %s,
                history = %s,
                diagnosis = %s,
                treatment = %s,
                examination = %s,
                finding = %s,
                advice = %s,
                nextAction = %s,
                username = %s,
                modified_date = %s
            WHERE id = %s
        """
        insert_sql = """
            INSERT INTO medic_report_dentalchart (
                parentid,
                toothid,
                toothplan,
                notationid,
                notationstatus,
                remarks,
                unitprice,
                billstatus,
                username,
                modified_date
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        """
        cursor = self._conn.cursor()
        try:
            cursor.execute(
                update_sql,
                (
                    generated_on,
                    appointment_on,
                    history,
                    diagnosis,
                    treatment,
                    examination,
                    finding,
                    advice,
                    next_action,
                    username,
                    datetime.now(),
                    int(report_id),
                ),
            )
            cursor.execute("DELETE FROM medic_report_dentalchart WHERE parentid = %s", (int(report_id),))
            if chart_items:
                payload = []
                timestamp = datetime.now()
                for item in chart_items:
                    payload.append(
                        (
                            int(report_id),
                            int(item.tooth_id),
                            (item.tooth_plan or "E")[:1],
                            int(item.notation_id),
                            int(item.notation_status),
                            item.remarks or "",
                            float(item.unit_price or 0.0),
                            int(item.bill_status),
                            username,
                            timestamp,
                        )
                    )
                cursor.executemany(insert_sql, payload)
            self._conn.commit()
        except Exception:
            self._conn.rollback()
            raise
        finally:
            cursor.close()

    def username_exists(self, username: str) -> bool:
        """Return True when a user record with the given username is present."""
        cleaned = (username or "").strip()
        if not cleaned:
            return False
        self._ensure_connection()
        sql = "SELECT 1 FROM users WHERE username = %s LIMIT 1"
        with self._conn.cursor() as cursor:
            cursor.execute(sql, (cleaned,))
            return cursor.fetchone() is not None

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            pass

