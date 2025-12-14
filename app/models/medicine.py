"""Dataclasses representing medicines and invoice items."""

from dataclasses import dataclass
from typing import Optional


@dataclass
class Medicine:
    name: str
    mg: str
    rack_number: str
    stock: int
    price: float
    actual_price: Optional[float] = None

    @property
    def effective_price(self) -> float:
        return self.actual_price if self.actual_price not in (None, "") else self.price


@dataclass
class InvoiceItem:
    medicine: Medicine
    quantity: int

    @property
    def line_total(self) -> float:
        return self.quantity * self.medicine.effective_price
