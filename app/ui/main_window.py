"""Main PyQt window for the Pharmacy Invoice Generator."""

from __future__ import annotations

from typing import Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QAbstractItemView,
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

from app.data.excel_repo import ExcelRepository
from app.models.medicine import InvoiceItem, Medicine
from app.printing.receipt_printer import ReceiptPrinter


class MainWindow(QMainWindow):
    """UI controller that ties together search, invoice, and printing."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Pharmacy Invoice Generator")
        self.setMinimumSize(900, 600)

        self.repo: Optional[ExcelRepository] = None
        self.medicines_by_name: Dict[str, Medicine] = {}
        self.invoice_items: List[InvoiceItem] = []
        self.invoice_quantities: Dict[str, int] = {}
        self.selected_medicine: Optional[Medicine] = None

        self._build_ui()
        self._load_inventory()

    def _build_ui(self) -> None:
        """Construct all widgets and layouts."""
        central = QWidget()
        root_layout = QVBoxLayout()

        # Search + details
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search medicine name...")
        self.search_input.textChanged.connect(self._on_search_text_changed)
        search_layout.addWidget(QLabel("Medicine:"))
        search_layout.addWidget(self.search_input, 1)

        details_layout = QFormLayout()
        self.mg_label = QLabel("-")
        self.rack_label = QLabel("-")
        self.stock_label = QLabel("-")
        self.price_label = QLabel("-")
        details_layout.addRow("MG:", self.mg_label)
        details_layout.addRow("Rack:", self.rack_label)
        details_layout.addRow("Stock:", self.stock_label)
        details_layout.addRow("Price:", self.price_label)

        qty_layout = QHBoxLayout()
        self.qty_spin = QSpinBox()
        self.qty_spin.setMinimum(1)
        self.qty_spin.setMaximum(100000)
        qty_layout.addWidget(QLabel("Qty:"))
        qty_layout.addWidget(self.qty_spin)

        self.add_button = QPushButton("Add to Invoice")
        self.add_button.clicked.connect(self._on_add_clicked)

        top_grid = QGridLayout()
        top_grid.addLayout(search_layout, 0, 0, 1, 2)
        top_grid.addLayout(details_layout, 1, 0)
        top_grid.addLayout(qty_layout, 1, 1)
        top_grid.addWidget(self.add_button, 2, 0, 1, 2)

        # Invoice table
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Medicine", "MG", "Rack", "Qty", "Price", "Total"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)

        # Totals
        totals_group = QGroupBox("Totals")
        totals_layout = QGridLayout()
        self.subtotal_value = QLabel("0.00")
        self.discount_spin = QDoubleSpinBox()
        self.discount_spin.setRange(0, 100)
        self.discount_spin.setDecimals(1)
        self.discount_spin.setSingleStep(0.5)
        self.discount_spin.valueChanged.connect(self._update_totals)
        self.net_total_value = QLabel("0.00")
        totals_layout.addWidget(QLabel("Subtotal"), 0, 0)
        totals_layout.addWidget(self.subtotal_value, 0, 1)
        totals_layout.addWidget(QLabel("Discount (%)"), 1, 0)
        totals_layout.addWidget(self.discount_spin, 1, 1)
        totals_layout.addWidget(QLabel("Net Total"), 2, 0)
        totals_layout.addWidget(self.net_total_value, 2, 1)
        totals_group.setLayout(totals_layout)

        # Print button
        self.print_button = QPushButton("PRINT")
        self.print_button.setStyleSheet("font-size: 16px; padding: 10px;")
        self.print_button.clicked.connect(self._on_print_clicked)

        root_layout.addLayout(top_grid)
        root_layout.addWidget(self.table, 1)
        root_layout.addWidget(totals_group)
        root_layout.addWidget(self.print_button, alignment=Qt.AlignRight)

        central.setLayout(root_layout)
        self.setCentralWidget(central)

    def _load_inventory(self) -> None:
        """Load medicines from Excel and configure completer."""
        self.repo = None
        try:
            self.repo = ExcelRepository()
            medicines = self.repo.list_medicines()
        except Exception as exc:  # noqa: BLE001 - surface Excel issues to user
            QMessageBox.critical(self, "Error", f"Failed to load Excel inventory:\n{exc}")
            medicines = []

        self.medicines_by_name = {m.name: m for m in medicines}
        completer = QCompleter(list(self.medicines_by_name.keys()))
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        completer.activated[str].connect(self._on_medicine_selected)
        self.search_input.setCompleter(completer)
        self._clear_selection()

    def _clear_selection(self) -> None:
        self.selected_medicine = None
        self.mg_label.setText("-")
        self.rack_label.setText("-")
        self.stock_label.setText("-")
        self.price_label.setText("-")

    def _on_search_text_changed(self, text: str) -> None:
        if text in self.medicines_by_name:
            self._on_medicine_selected(text)

    def _on_medicine_selected(self, name: str) -> None:
        medicine = self.medicines_by_name.get(name)
        if not medicine:
            self._clear_selection()
            return
        self.selected_medicine = medicine
        self.mg_label.setText(medicine.mg)
        self.rack_label.setText(medicine.rack_number)
        self.stock_label.setText(str(self._available_stock(medicine)))
        self.price_label.setText(f"{medicine.effective_price:.2f}")

    def _available_stock(self, medicine: Medicine) -> int:
        """Return stock considering items already added to invoice."""
        used = self.invoice_quantities.get(medicine.name, 0)
        return max(medicine.stock - used, 0)

    def _on_add_clicked(self) -> None:
        if not self.selected_medicine:
            QMessageBox.warning(self, "Select Medicine", "Please choose a medicine first.")
            return
        qty = int(self.qty_spin.value())
        available = self._available_stock(self.selected_medicine)
        if qty > available:
            QMessageBox.warning(
                self,
                "Insufficient Stock",
                f"Only {available} units available for {self.selected_medicine.name}.",
            )
            return

        self._add_or_update_item(self.selected_medicine, qty)
        self._refresh_table()
        self._update_totals()
        self._on_medicine_selected(self.selected_medicine.name)

    def _add_or_update_item(self, medicine: Medicine, qty: int) -> None:
        """Add new item or increase quantity if already present."""
        for item in self.invoice_items:
            if item.medicine.name == medicine.name:
                item.quantity += qty
                break
        else:
            self.invoice_items.append(InvoiceItem(medicine=medicine, quantity=qty))

        self.invoice_quantities[medicine.name] = self.invoice_quantities.get(medicine.name, 0) + qty

    def _refresh_table(self) -> None:
        self.table.setRowCount(len(self.invoice_items))
        for row, item in enumerate(self.invoice_items):
            values = [
                item.medicine.name,
                item.medicine.mg,
                item.medicine.rack_number,
                str(item.quantity),
                f"{item.medicine.effective_price:.2f}",
                f"{item.line_total:.2f}",
            ]
            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))
        self.table.resizeColumnsToContents()

    def _update_totals(self) -> None:
        subtotal = sum(item.line_total for item in self.invoice_items)
        discount_pct = float(self.discount_spin.value())
        net_total = subtotal * (1 - discount_pct / 100)
        self.subtotal_value.setText(f"{subtotal:.2f}")
        self.net_total_value.setText(f"{net_total:.2f}")

    def _on_print_clicked(self) -> None:
        if not self.invoice_items:
            QMessageBox.information(self, "Nothing to print", "Add at least one item to the invoice.")
            return

        subtotal = sum(item.line_total for item in self.invoice_items)
        discount_pct = float(self.discount_spin.value())
        printer = ReceiptPrinter()
        success = printer.print_receipt(self.invoice_items, subtotal, discount_pct)

        if not success:
            QMessageBox.critical(self, "Print Failed", "Printer is not available or failed to print.")
            return

        try:
            self._persist_stock_changes()
            QMessageBox.information(self, "Printed", "Receipt sent to printer and stock updated.")
            self._reset_invoice()
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Excel Error", f"Could not update Excel stock:\n{exc}")

    def _persist_stock_changes(self) -> None:
        """Write new stock levels back to Excel."""
        if not self.repo:
            raise RuntimeError("Inventory repository is not initialized.")
        new_levels: Dict[str, int] = {}
        for med in self.medicines_by_name.values():
            used = self.invoice_quantities.get(med.name, 0)
            new_levels[med.name] = max(med.stock - used, 0)
        self.repo.save_stock_updates(new_levels)

    def _reset_invoice(self) -> None:
        self.invoice_items.clear()
        self.invoice_quantities.clear()
        self.table.setRowCount(0)
        self.discount_spin.setValue(0.0)
        self._update_totals()
        self._load_inventory()
