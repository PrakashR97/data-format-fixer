"""Creates samples/sample_export.xlsx - a messy 'database export' for trying the app."""
from pathlib import Path
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Orders"
ws.append(["Order ID", "Customer", "Quantity", "Unit Price", "Order Date", "Shipped At", "Phone", "Zip", "Notes", "Total"])
rows = [
    ("1001", "Ana Silva", "3", "19.99", "2026-01-15", "2026-01-16 09:30:00", "5551234567", "02134", "Gift wrap", None),
    ("1002", "Ben Carter", "12", "5.50", "2026-02-03", "2026-02-04 14:05:10", "5559876543", "90210", "", None),
    ("1003", "Chloe Dubois", "1,200", "0.75", "2026-02-28", "2026-03-01 08:00:00", "5550001111", "10001", "Bulk", None),
    ("1004", "Dev Patel", "7", "102.00", "2026-03-10", "2026-03-11 17:45:00", "5552223333", "00501", "N/A", None),
    ("1005", "Eli Novak", "n/a", "8.25", "2026-03-31", "", "5554445555", "30301", "Check address", None),
    ("1006", "Fay Chen", "4", "abc", "2026-02-30", "2026-04-02 10:10:10", "5556667777", "60601", "Bad price + impossible date", None),
]
for i, r in enumerate(rows, start=2):
    ws.append(list(r))
    ws[f"J{i}"] = f"=C{i}*D{i}"  # formula cell - must stay untouched (note: C/D are text here)

dates = wb.create_sheet("Dates")
dates.append(["Ambiguous", "Day-first clue", "Written out", "Long ID", "Mixed column"])
for row in [
    ("01/02/2026", "13/02/2026", "5 March 2026", "12345678901234567890", "42"),
    ("03/04/2026", "01/02/2026", "Mar 6, 2026", "98765432109876543210", "hello"),
    ("05/06/2026", "20/02/2026", "7 Mar 2026 2:30 PM", "11112222333344445555", "2026-01-01"),
    ("07/08/2026", "02/03/2026", "March 8, 2026", "55556666777788889999", "17"),
]:
    dates.append(list(row))

amounts = wb.create_sheet("Amounts")
amounts.append(["Revenue", "Count", "Rate"])
for r in [("1,234.50", "10", "0.125"), ("-980.00", "20", "0.5"), ("45,000.25", "30", "1.75"), ("", "40", "")]:
    amounts.append(list(r))

out = Path(__file__).parent / "sample_export.xlsx"
wb.save(out)
print("Wrote", out)
