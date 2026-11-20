"""Regenerates the .xlsx fixtures used by tests (needs: pip install openpyxl xlsxwriter; LibreOffice for the third file).
The fixtures are committed, so the tests do not need Python."""
import datetime, subprocess, sys, os, tempfile, shutil

out = os.path.dirname(os.path.abspath(__file__))

# 1) openpyxl -------------------------------------------------------------
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Cars"
ws.append(["الماركة", "الموديل", "Year", "Price", "Available", "Added", "Updated"])
ws.append(["هيونداي", "كريتا", 2021, 15999.5, True, datetime.date(2024, 3, 5), datetime.datetime(2024, 3, 5, 13, 45, 30)])
ws.append(["Toyota", "Corolla & Co <x> \"q\"", 2020, None, False, None, None])
ws.append(["", None, None, None, None, None, None])
ws.append(["Zero", "00123", -7, 0.1, None, None, None])
ws.append(["Multi", "line1\nline2", 1, 2, 3, 4, 5])
ws2 = wb.create_sheet("ملاحظات")
ws2["B3"] = "offset cell"
ws2["D5"] = 42
wb.save(os.path.join(out, "openpyxl.xlsx"))

# 2) xlsxwriter -----------------------------------------------------------
import xlsxwriter
wb = xlsxwriter.Workbook(os.path.join(out, "xlsxwriter.xlsx"))
ws = wb.add_worksheet("Mixed")
bold = wb.add_format({"bold": True})
pct = wb.add_format({"num_format": "0.0%"})
money = wb.add_format({"num_format": '#,##0.00 "SAR"'})
date = wb.add_format({"num_format": "dd/mm/yyyy"})
ws.write_rich_string("A1", bold, "Rich ", "text")
ws.write("A2", "plain")
ws.write_number("B1", 0.25, pct)
ws.write_number("B2", 1234.5, money)
ws.write_formula("C1", "=B1*2", None, 0.5)
ws.write_formula("C2", '="a"&"b"', None, "ab")
ws.write_datetime("D1", datetime.datetime(2000, 1, 31), date)
ws.write_boolean("D2", True)
ws.write_string("E1", "  spaces  ")
ws.write_blank("F1", None, bold)
wb.close()

# 3) LibreOffice (a third producer) ---------------------------------------
tmp = tempfile.mkdtemp()
csv_path = os.path.join(tmp, "libreoffice.csv")
with open(csv_path, "w", encoding="utf-8") as f:
    f.write('name,qty,note\nشاي,3,"a, b"\nقهوة,12.5,x\n')
try:
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, csv_path],
                   check=True, capture_output=True, timeout=120)
except Exception as e:
    print("LibreOffice step skipped:", e, file=sys.stderr)
shutil.rmtree(tmp, ignore_errors=True)
print("fixtures written to", out)
