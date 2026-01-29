"""Generate printable receipt PDFs."""
from __future__ import annotations

from dataclasses import dataclass
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

split_lines = _split_lines

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

