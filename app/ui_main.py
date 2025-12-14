"""PyQt5 UI for the Pharmacy Invoice Generator."""

from __future__ import annotations

from typing import Dict, Optional

from PyQt5.QtCore import Qt, QStringListModel
from PyQt5.QtGui import QFont
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtGui import QTextDocument
from PyQt5.QtWidgets import (
    QApplication,
    QCompleter,
    QDoubleSpinBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app import config
from app.data_store import ExcelDataStore
from app.models import Invoice, InvoiceItem, format_currency


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Pharmacy Invoice Generator")
        self.resize(1000, 600)

        self.store: Optional[ExcelDataStore] = None
        self.invoice = Invoice()
        self.added_qty: Dict[str, int] = {}
        self.selected_name: Optional[str] = None
        self._name_model = QStringListModel()
        self._completer: Optional[QCompleter] = None

        self._build_ui()
        self._load_store()

    def _build_ui(self) -> None:
        root = QWidget()
        main_layout = QVBoxLayout()

        content_layout = QHBoxLayout()
        content_layout.addLayout(self._build_left_panel(), 1)
        content_layout.addLayout(self._build_right_panel(), 2)

        main_layout.addLayout(content_layout)
        main_layout.addLayout(self._build_bottom_panel())
        root.setLayout(main_layout)
        self.setCentralWidget(root)

    def _build_left_panel(self) -> QVBoxLayout:
        layout = QVBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search medicine...")
        self.search_input.textChanged.connect(self._on_search_text)
        self.search_input.returnPressed.connect(self._on_search_enter)
        layout.addWidget(QLabel("Medicine"))
        layout.addWidget(self.search_input)

        details = QFormLayout()
        self.mg_label = QLabel("-")
        self.rack_label = QLabel("-")
        self.stock_label = QLabel("-")
        self.price_label = QLabel("-")
        details.addRow("MG:", self.mg_label)
        details.addRow("Rack:", self.rack_label)
        details.addRow("Stock:", self.stock_label)
        details.addRow("Price:", self.price_label)
        layout.addLayout(details)

        qty_layout = QHBoxLayout()
        qty_layout.addWidget(QLabel("Qty:"))
        self.qty_spin = QSpinBox()
        self.qty_spin.setMinimum(1)
        self.qty_spin.setMaximum(1_000_000)
        qty_layout.addWidget(self.qty_spin)
        layout.addLayout(qty_layout)

        self.add_button = QPushButton("Add to Invoice")
        self.add_button.clicked.connect(self._add_to_invoice)
        layout.addWidget(self.add_button)

        layout.addStretch()
        return layout

    def _build_right_panel(self) -> QVBoxLayout:
        layout = QVBoxLayout()
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Medicine", "MG", "Qty", "Price", "Total"])
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table, 1)
        return layout

    def _build_bottom_panel(self) -> QHBoxLayout:
        bottom = QHBoxLayout()

        info_group = QGroupBox("Customer")
        info_layout = QFormLayout()
        self.customer_name = QLineEdit()
        self.customer_phone = QLineEdit()
        info_layout.addRow("Name", self.customer_name)
        info_layout.addRow("Phone", self.customer_phone)
        info_group.setLayout(info_layout)

        totals_group = QGroupBox("Totals")
        totals_layout = QGridLayout()
        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setRange(0, 100)
        self.discount_spin.setDecimals(1)
        self.discount_spin.valueChanged.connect(self._update_totals)

        self.total_label = QLabel("0.00")
        total_font = QFont()
        total_font.setPointSize(16)
        total_font.setBold(True)
        self.total_label.setFont(total_font)

        totals_layout.addWidget(QLabel("Discount %"), 0, 0)
        totals_layout.addWidget(self.discount_spin, 0, 1)
        totals_layout.addWidget(QLabel("Net Total"), 1, 0)
        totals_layout.addWidget(self.total_label, 1, 1)
        totals_group.setLayout(totals_layout)

        buttons_layout = QVBoxLayout()
        self.print_button = QPushButton("PRINT")
        self.print_button.clicked.connect(self._on_print)
        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self._clear_invoice)
        buttons_layout.addWidget(self.print_button)
        buttons_layout.addWidget(self.clear_button)
        buttons_layout.addStretch()

        bottom.addWidget(info_group)
        bottom.addWidget(totals_group)
        bottom.addLayout(buttons_layout)
        return bottom

    def _load_store(self) -> None:
        try:
            self.store = ExcelDataStore()
            if not self.store.self_check():
                raise ValueError("Excel sheet missing required columns.")
            names = self.store.get_all_names()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Excel Error", f"Failed to load Excel: {exc}")
            self.add_button.setEnabled(False)
            self.print_button.setEnabled(False)
            names = []

        self._name_model.setStringList(names)
        completer = QCompleter(self._name_model, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        completer.activated[str].connect(self._select_medicine)
        self._completer = completer
        self.search_input.setCompleter(completer)

    def _on_search_text(self, text: str) -> None:
        if not text:
            self._clear_selection()
            return
        if not self.store:
            return
        if self._completer:
            self._completer.setCompletionPrefix(text)
            self._completer.complete()
        # Try exact match first (case-insensitive)
        record = self.store.get_medicine(text)
        if record:
            self._select_medicine(record["name"])

    def _on_search_enter(self) -> None:
        """Handle Enter key to select medicine or show not-found."""
        text = self.search_input.text().strip()
        if not text or not self.store:
            return
        record = self.store.get_medicine(text)
        if record:
            self._select_medicine(record["name"])
        else:
            QMessageBox.information(self, "Not found", "Medicine not found in Excel.")
            self._clear_selection()

    def _select_medicine(self, name: str) -> None:
        if not self.store:
            return
        record = self.store.get_medicine(name)
        if not record:
            self._clear_selection()
            return
        self.selected_name = record["name"]
        self.mg_label.setText(record["mg"])
        self.rack_label.setText(record["rack"])
        self.stock_label.setText(str(self._available_stock(record["name"])))
        self.price_label.setText(format_currency(record["price"]))

    def _clear_selection(self) -> None:
        self.selected_name = None
        self.mg_label.setText("-")
        self.rack_label.setText("-")
        self.stock_label.setText("-")
        self.price_label.setText("-")

    def _available_stock(self, name: str) -> int:
        if not self.store:
            return 0
        record = self.store.get_medicine(name)
        if not record:
            return 0
        used = self.added_qty.get(record["name"], 0)
        return max(record["stock"] - used, 0)

    def _add_to_invoice(self) -> None:
        if not self.store or not self.selected_name:
            QMessageBox.warning(self, "Select medicine", "Please select a medicine first.")
            return
        record = self.store.get_medicine(self.selected_name)
        if not record:
            QMessageBox.warning(self, "Missing", "Selected medicine not found.")
            return

        qty = int(self.qty_spin.value())
        available = self._available_stock(record["name"])
        if qty > available:
            QMessageBox.warning(
                self,
                "Insufficient stock",
                f"Only {available} units available for {record['name']}.",
            )
            return

        item = InvoiceItem(
            name=record["name"],
            mg=record["mg"],
            rack=record["rack"],
            unit_price=record["price"],
            qty=qty,
        )
        self.invoice.add_item(item)
        self.added_qty[record["name"]] = self.added_qty.get(record["name"], 0) + qty
        self._refresh_table()
        self._update_totals()
        self._select_medicine(record["name"])

    def _refresh_table(self) -> None:
        self.table.setRowCount(len(self.invoice.items))
        for row, item in enumerate(self.invoice.items):
            values = [
                item.name,
                item.mg,
                str(item.qty),
                format_currency(item.unit_price),
                format_currency(item.line_total),
            ]
            for col, val in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(val))
        self.table.resizeColumnsToContents()

    def _update_totals(self) -> None:
        net = self.invoice.net_total(self.discount_spin.value())
        self.total_label.setText(format_currency(net))

    def _build_receipt_html(self) -> str:
        rows = []
        for item in self.invoice.items:
            rows.append(
                f"<tr><td>{item.name}</td>"
                f"<td align='right'>{item.qty}</td>"
                f"<td align='right'>{item.unit_price:.2f}</td>"
                f"<td align='right'>{item.line_total:.2f}</td></tr>"
            )
        subtotal = self.invoice.subtotal()
        discount_pct = self.discount_spin.value()
        discount_amt = self.invoice.discount_amount(discount_pct)
        net = self.invoice.net_total(discount_pct)
        customer = self.customer_name.text().strip()
        phone = self.customer_phone.text().strip()
        customer_line = customer or "Walk-in Customer"
        if phone:
            customer_line += f" ({phone})"

        return f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial; font-size: 11pt; }}
                h2 {{ text-align: center; margin: 0 0 8px 0; }}
                table {{ width: 100%; border-collapse: collapse; }}
                th, td {{ padding: 2px 0; }}
                .totals td {{ padding-top: 4px; }}
            </style>
        </head>
        <body>
            <h2>{config.STORE_HEADER}</h2>
            <p>{customer_line}</p>
            <table>
                <tr><th align='left'>Item</th><th align='right'>Qty</th><th align='right'>Rate</th><th align='right'>Total</th></tr>
                {''.join(rows)}
            </table>
            <hr />
            <table class='totals'>
                <tr><td>Subtotal</td><td align='right'>{subtotal:.2f}</td></tr>
                <tr><td>Discount ({discount_pct:.1f}%)</td><td align='right'>-{discount_amt:.2f}</td></tr>
                <tr><td><b>Net Total</b></td><td align='right'><b>{net:.2f}</b></td></tr>
            </table>
        </body>
        </html>
        """

    def _on_print(self) -> None:
        if not self.store:
            QMessageBox.warning(self, "Excel not loaded", "Cannot print without Excel data.")
            return
        if not self.invoice.items:
            QMessageBox.information(self, "No items", "Add at least one item to print.")
            return

        printer = QPrinter(QPrinter.HighResolution)
        printer.setPrinterName(config.PRINTER_NAME)
        if not printer.isValid():
            QMessageBox.critical(self, "Printer Error", "Printer not available.")
            return

        doc = QTextDocument()
        doc.setHtml(self._build_receipt_html())
        doc.print_(printer)

        try:
            for item in self.invoice.items:
                self.store.reduce_stock(item.name, item.qty)
            self.store.save()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Excel Error", f"Failed to update stock: {exc}")
            return

        QMessageBox.information(self, "Printed", "Receipt sent to printer and stock updated.")
        self._clear_invoice()

    def _clear_invoice(self) -> None:
        self.invoice = Invoice()
        self.added_qty.clear()
        self.table.setRowCount(0)
        self.discount_spin.setValue(0.0)
        self._update_totals()
        self._clear_selection()
        if self.store:
            self._load_store()


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()
