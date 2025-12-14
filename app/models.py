"""Invoice data models."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass
class InvoiceItem:
    name: str
    mg: str
    rack: str
    unit_price: float
    qty: int

    @property
    def line_total(self) -> float:
        return self.unit_price * self.qty


@dataclass
class Invoice:
    items: List[InvoiceItem] = field(default_factory=list)

    def add_item(self, item: InvoiceItem) -> None:
        """Add or merge an item by medicine name."""
        for existing in self.items:
            if existing.name == item.name:
                existing.qty += item.qty
                return
        self.items.append(item)

    def subtotal(self) -> float:
        return sum(item.line_total for item in self.items)

    def discount_amount(self, percent: float) -> float:
        return self.subtotal() * (percent / 100.0)

    def net_total(self, percent: float) -> float:
        return self.subtotal() - self.discount_amount(percent)

    def total_qty(self) -> int:
        return sum(item.qty for item in self.items)


def format_currency(amount: float) -> str:
    """Return amount formatted to two decimals."""
    return f"{amount:.2f}"

