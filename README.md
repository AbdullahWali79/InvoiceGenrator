# Pharmacy Invoice Generator (PyQt5)

Desktop app to create and print thermal receipts for pharmacy sales using an Excel workbook as the only data source.

## Prerequisites
- Python 3.10+ on Windows
- Excel workbook at `data/medicines.xlsx` with sheet `Medicines` and headers in row 1:
  - `Medicine_Name`, `MG`, `Rack_Number`, `Stock`, `Price`, `Actual_Price`

## Setup
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Run
```powershell
python main.py
```

## Usage Notes
- Start typing a medicine name to get autosuggestions from Excel.
- Price uses `Actual_Price` when present, otherwise `Price`.
- Quantity cannot exceed available stock (including items already in the invoice).
- After printing, stock is reduced in Excel and the workbook is saved.
- Default printer name is `Star TSP700II (TSP743II)`; change it in `app/config.py` if needed.
