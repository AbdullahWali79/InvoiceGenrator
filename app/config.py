"""Configuration constants for the Pharmacy Invoice Generator."""

from pathlib import Path

# Path to the Excel workbook containing the inventory data.
EXCEL_PATH: Path = Path("data/medicines.xlsx")

# Sheet name inside the Excel workbook.
EXCEL_SHEET_NAME: str = "Medicines"

# Name of the Windows printer to target for receipts.
PRINTER_NAME: str = "Star TSP700II (TSP743II)"

# Receipt paper width in millimeters for 80mm thermal rolls.
RECEIPT_WIDTH_MM: float = 80.0

# Header text that prints on top of the receipt.
STORE_HEADER: str = "Pharmacy Invoice"
