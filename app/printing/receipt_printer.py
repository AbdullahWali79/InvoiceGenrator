"""Receipt printing via QTextDocument to a Windows printer."""

from __future__ import annotations

from typing import Iterable, List

from PyQt5.QtGui import QTextDocument
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtCore import QSizeF

from app import config
from app.models.medicine import InvoiceItem


class ReceiptPrinter:
    """Render and print receipts as HTML to a target printer."""

    def __init__(self, printer_name: str | None = None, receipt_width_mm: float | None = None) -> None:
        self.printer_name = printer_name or config.PRINTER_NAME
        self.receipt_width_mm = receipt_width_mm or config.RECEIPT_WIDTH_MM

    def _build_html(self, items: Iterable[InvoiceItem], subtotal: float, discount_pct: float) -> str:
        rows: List[str] = []
        for item in items:
            rows.append(
                f"<tr><td>{item.medicine.name}</td>"
                f"<td align='right'>{item.quantity}</td>"
                f"<td align='right'>{item.medicine.effective_price:.2f}</td>"
                f"<td align='right'>{item.line_total:.2f}</td></tr>"
            )

        discount_amount = subtotal * (discount_pct / 100)
        net_total = subtotal - discount_amount

        return f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Arial'; font-size: 11pt; }}
                h2 {{ text-align: center; margin: 0 0 6px 0; }}
                table {{ width: 100%; border-collapse: collapse; }}
                td {{ padding: 2px 0; }}
                .totals td {{ padding-top: 4px; }}
            </style>
        </head>
        <body>
            <h2>{config.STORE_HEADER}</h2>
            <table>
                <tr><th align='left'>Item</th><th align='right'>Qty</th><th align='right'>Rate</th><th align='right'>Total</th></tr>
                {''.join(rows)}
            </table>
            <hr />
            <table class='totals'>
                <tr><td>Subtotal</td><td align='right'>{subtotal:.2f}</td></tr>
                <tr><td>Discount ({discount_pct:.1f}%)</td><td align='right'>-{discount_amount:.2f}</td></tr>
                <tr><td><b>Net Total</b></td><td align='right'><b>{net_total:.2f}</b></td></tr>
            </table>
            <p style='text-align:center;margin-top:8px;'>Thank you!</p>
        </body>
        </html>
        """

    def print_receipt(self, items: Iterable[InvoiceItem], subtotal: float, discount_pct: float) -> bool:
        """Send receipt to printer; returns True on success."""
        items_list = list(items)
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPrinterName(self.printer_name)

        if not printer.isValid():
            return False

        # Dynamic height to avoid truncation; 40mm base plus 8mm per line.
        height_mm = 40 + (len(items_list) * 8)
        printer.setPaperSize(QSizeF(self.receipt_width_mm, height_mm), QPrinter.Millimeter)
        printer.setFullPage(True)

        doc = QTextDocument()
        doc.setHtml(self._build_html(items_list, subtotal, discount_pct))
        doc.setPageSize(QSizeF(self.receipt_width_mm, height_mm))

        doc.print_(printer)
        return printer.isValid()
