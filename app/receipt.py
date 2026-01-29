"""Generate printable receipt PDFs."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import re
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

from .config import ClinicInfo
from .database import Patient, Receipt, ReceiptItem


@dataclass
class PaymentEntry:
    method: str
    amount: float


@dataclass
class PaymentProgress:
    sequence: int
    current_amount: float
    total_paid: float
    balance: float
    total_due: float
    received_on: datetime
    method: str
    previous_payments: Sequence[Tuple[int, datetime, float]] = ()
    remark: str = ""


class ReceiptPDFGenerator:
    """Lay out receipt data into a polished PDF document."""

    def __init__(self, output_dir: Path, currency: str = "RM") -> None:
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.currency = currency

    def generate(
        self,
        clinic: ClinicInfo,
        patient: Patient,
        receipt: Receipt,
        items: Sequence[ReceiptItem],
        payments: Sequence[PaymentEntry],
        logo_path: Optional[Path] = None,
        payment_progress: Optional[PaymentProgress] = None,
    ) -> Path:
        file_name = self._build_filename(patient, receipt, payment_progress)
        output_path = self.output_dir / file_name

        pdf = canvas.Canvas(str(output_path), pagesize=A4)
        width, height = A4
        margin = 18 * mm
        content_top = height - margin

        header_bottom = self._draw_header(pdf, clinic, logo_path, margin, content_top)
        current_y = header_bottom - 14
        self._draw_title(pdf, "Official Receipt", width / 2, current_y)
        current_y -= 18
        pdf.line(margin, current_y, width - margin, current_y)
        current_y -= 20

        current_y = self._draw_patient_block(pdf, clinic, patient, receipt, margin, width - margin, current_y)
        current_y -= 22

        current_y = self._draw_items_table(pdf, items, margin, width - margin, current_y)
        current_y -= 24

        self._draw_totals(pdf, receipt, payments, margin, width - margin, current_y, payment_progress)

        pdf.showPage()
        pdf.save()
        return output_path

    # ------------------------------------------------------------------

    def _build_filename(
        self,
        patient: Patient,
        receipt: Receipt,
        payment_progress: Optional[PaymentProgress],
    ) -> str:
        patient_label = (
            patient.receipt_name
            or patient.preferred_name
            or patient.name
            or patient.icpassport
            or "patient"
        )
        receipt_label = receipt.rcpt_id or "receipt"
        patient_slug = self._slugify(patient_label)
        receipt_slug = self._slugify(receipt_label)
        if not patient_slug:
            patient_slug = "patient"
        if not receipt_slug:
            receipt_slug = "receipt"

        date_source = payment_progress.received_on if payment_progress else receipt.issued
        date_label = date_source.strftime("%Y%m%d")
        seq_suffix = ""
        if payment_progress and payment_progress.sequence:
            seq_suffix = f"_p{payment_progress.sequence:02d}"
        return f"receipt_{patient_slug}_{receipt_slug}_{date_label}{seq_suffix}.pdf"

    @staticmethod
    def _slugify(text: str) -> str:
        cleaned = re.sub(r"[^A-Za-z0-9]+", "_", text.strip())
        return cleaned.strip("_")[:80]

    def _draw_header(
        self,
        pdf: canvas.Canvas,
        clinic: ClinicInfo,
        logo_path: Optional[Path],
        left: float,
        top: float,
    ) -> float:
        text_x = left
        text_top = top - 6

        if logo_path and Path(logo_path).exists():
            try:
                image = ImageReader(str(logo_path))
                img_w, img_h = image.getSize()
                target_w = 52 * mm
                target_h = 30 * mm
                scale = min(target_w / img_w, target_h / img_h, 1.0)
                draw_w = img_w * scale
                draw_h = img_h * scale
                pdf.drawImage(
                    image,
                    left,
                    top - draw_h,
                    width=draw_w,
                    height=draw_h,
                    mask='auto',
                    preserveAspectRatio=True,
                )
                text_x += draw_w + 10
                text_top = min(text_top, top - draw_h + draw_h - 10)
            except Exception:
                pass

        text = pdf.beginText()
        text.setTextOrigin(text_x, top - 6)
        text.setFont("Helvetica-Bold", 14)
        text.textLine(clinic.name.strip() or "Clinic")
        text.setFont("Helvetica", 10)
        max_text_width = max(10.0, pdf._pagesize[0] - text_x - left)
        for raw_line in self._split_lines(clinic.address):
            for wrapped_line in self._wrap_text(raw_line, max_text_width, "Helvetica", 10):
                text.textLine(wrapped_line)
        if clinic.phone:
            text.textLine(f"Phone: {clinic.phone}")
        if clinic.email:
            text.textLine(f"Email: {clinic.email}")
        pdf.drawText(text)
        return text.getY()

    def _draw_title(self, pdf: canvas.Canvas, title: str, x: float, y: float) -> None:
        pdf.setFont("Helvetica-Bold", 18)
        pdf.drawCentredString(x, y, title)

    def _draw_patient_block(
        self,
        pdf: canvas.Canvas,
        clinic: ClinicInfo,
        patient: Patient,
        receipt: Receipt,
        left: float,
        right: float,
        top: float,
    ) -> float:
        column_gap = 14
        column_width = (right - left - column_gap) / 2

        patient_text = pdf.beginText()
        patient_text.setTextOrigin(left, top)
        patient_text.setFont("Helvetica-Bold", 11)
        patient_text.textLine("Received from")
        patient_text.setFont("Helvetica", 10)
        patient_text.textLine(patient.receipt_name.strip() or patient.name)
        patient_text.textLine(f"IC / Passport: {patient.icpassport}")
        if patient.company.strip():
            patient_text.textLine(f"Company: {patient.company.strip()}")
        if patient.phone_mobile.strip():
            patient_text.textLine(f"Mobile: {patient.phone_mobile.strip()}")
        elif patient.phone_fixed.strip():
            patient_text.textLine(f"Phone: {patient.phone_fixed.strip()}")
        if patient.email.strip():
            patient_text.textLine(f"Email: {patient.email.strip()}")
        pdf.drawText(patient_text)

        meta_text = pdf.beginText()
        meta_text.setTextOrigin(left + column_width + column_gap, top)
        meta_text.setFont("Helvetica", 10)
        meta_text.textLine(f"Receipt No: {receipt.rcpt_id}")
        meta_text.textLine(f"Date: {receipt.issued.strftime('%Y-%m-%d %H:%M')}")
        if receipt.done_by.strip():
            meta_text.textLine(f"Processed by: {receipt.done_by.strip()}")
        if receipt.department_type.strip():
            meta_text.textLine(f"Department: {receipt.department_type.strip()}")
        if receipt.settled_by.strip():
            meta_text.textLine(f"Settled by: {receipt.settled_by.strip()}")
        pdf.drawText(meta_text)

        return min(patient_text.getY(), meta_text.getY())

    def _draw_items_table(
        self,
        pdf: canvas.Canvas,
        items: Sequence[ReceiptItem],
        left: float,
        right: float,
        top: float,
    ) -> float:
        width = right - left
        column_ratios = [0.50, 0.12, 0.18, 0.20]
        col_widths = [width * ratio for ratio in column_ratios]

        header_height = 18
        pdf.setFillColorRGB(0.9, 0.9, 0.9)
        pdf.rect(left, top - header_height, width, header_height, fill=1, stroke=0)
        pdf.setFillColorRGB(0, 0, 0)
        pdf.rect(left, top - header_height, width, header_height, fill=0, stroke=1)

        pdf.setFont("Helvetica-Bold", 10)
        headers = ["Description", "Qty", "Unit Price", "Amount"]
        for idx, header in enumerate(headers):
            column_left = left + sum(col_widths[:idx])
            column_right = column_left + col_widths[idx]
            if idx == 0:
                pdf.drawString(column_left + 6, top - header_height + 5, header)
            else:
                pdf.drawRightString(column_right - 6, top - header_height + 5, header)

        current_y = top - header_height
        pdf.setFont("Helvetica", 9)
        for item in items:
            description = item.name or item.item_id or ""
            wrapped = self._wrap_text(description, col_widths[0] - 10, "Helvetica", 9)
            line_count = max(1, len(wrapped))
            row_height = 8 + line_count * 11
            current_y -= row_height
            pdf.rect(left, current_y, width, row_height, fill=0, stroke=1)

            text_y = current_y + row_height - 12
            for line in wrapped:
                pdf.drawString(left + 6, text_y, line)
                text_y -= 11

            qty = str(item.qty)
            unit = self._fmt_currency(item.unit_price)
            amount_value = (item.subtotal or 0.0) - (item.discount or 0.0)
            amount = self._fmt_currency(amount_value)

            pdf.drawRightString(left + col_widths[0] + col_widths[1] - 6, current_y + 6, qty)
            pdf.drawRightString(left + col_widths[0] + col_widths[1] + col_widths[2] - 6, current_y + 6, unit)
            pdf.drawRightString(right - 6, current_y + 6, amount)

        return current_y

    def _draw_totals(
        self,
        pdf: canvas.Canvas,
        receipt: Receipt,
        payments: Sequence[PaymentEntry],
        left: float,
        right: float,
        top: float,
        payment_progress: Optional[PaymentProgress],
    ) -> None:
        width = right - left
        summary_width = width * 0.42
        start_x = right - summary_width
        line_y = top

        subtotal_value = receipt.subtotal or 0.0
        total_due = receipt.total or 0.0
        if payment_progress:
            previous_paid = max(payment_progress.total_paid - payment_progress.current_amount, 0.0)
            current_payment = payment_progress.current_amount
            balance_payment = max(payment_progress.balance, 0.0)
            total_due = payment_progress.total_due
        else:
            previous_paid = 0.0
            current_payment = total_due
            balance_payment = max(total_due - current_payment, 0.0)

        is_installment = (
            payment_progress is not None
            and (
                payment_progress.previous_payments
                or balance_payment > 0.01
                or abs(current_payment - total_due) > 0.01
            )
        )

        if is_installment:
            summary_rows = [
                ("Subtotal", subtotal_value),
                ("Previous Payments", previous_paid),
                ("Current Payment", current_payment),
                ("Balance Payment", balance_payment),
                ("Total", total_due),
            ]
        else:
            summary_rows = [
                ("Subtotal", subtotal_value),
                ("Total", total_due),
            ]

        for label, amount in summary_rows:
            if label == "Total":
                pdf.setFont("Helvetica-Bold", 11)
            else:
                pdf.setFont("Helvetica", 10)
            pdf.drawString(start_x, line_y, label)
            pdf.drawRightString(right, line_y, self._fmt_currency(amount))
            line_y -= 16

        line_y -= 6
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(left, line_y, "Payment Details")
        line_y -= 16
        pdf.setFont("Helvetica", 10)
        if payments:
            for entry in payments:
                pdf.drawString(left, line_y, entry.method)
                pdf.drawRightString(right, line_y, self._fmt_currency(entry.amount))
                line_y -= 14
        else:
            pdf.drawString(left, line_y, "Payment data unavailable")
            line_y -= 14

        if not is_installment:
            return

        line_y -= 4
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(left, line_y, "Payment Progress")
        line_y -= 16
        pdf.setFont("Helvetica", 10)

        previous_payments = payment_progress.previous_payments if payment_progress else ()
        if previous_payments:
            for seq, paid_on, amount in previous_payments:
                label = f"Payment #{seq} on {paid_on.strftime('%d-%m-%Y')}"
                pdf.drawString(left, line_y, label)
                pdf.drawRightString(right, line_y, self._fmt_currency(amount))
                line_y -= 14
        else:
            pdf.drawString(left, line_y, "No previous payments")
            line_y -= 14

        balance_display = payment_progress.balance if payment_progress else balance_payment
        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(left, line_y, "Balance Remaining")
        pdf.drawRightString(right, line_y, self._fmt_currency(balance_display))

    # ------------------------------------------------------------------

    def _fmt_currency(self, value: float) -> str:
        return f"{self.currency} {value:,.2f}"

    def _split_lines(self, text: str) -> Iterable[str]:
        if not text:
            return []
        lines = []
        for segment in text.replace('\r', '').split('\n'):
            cleaned = segment.strip()
            if cleaned:
                lines.append(cleaned)
        return lines

    def _wrap_text(self, text: str, max_width: float, font: str, size: int) -> List[str]:
        words = text.split()
        if not words:
            return [""]
        lines: List[str] = []
        current = words[0]
        for word in words[1:]:
            proposal = f"{current} {word}"
            if stringWidth(proposal, font, size) <= max_width:
                current = proposal
            else:
                lines.append(current)
                current = word
        lines.append(current)
        return lines


# provide backward-compatible alias
ReceiptPDFGenerator.split_lines = ReceiptPDFGenerator._split_lines

