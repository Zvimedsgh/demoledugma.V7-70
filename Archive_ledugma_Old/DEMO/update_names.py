import openpyxl

for year in ["2024", "2025"]:
    wb = openpyxl.load_workbook(fr"C:\ledugma\DEMO\{year}.xlsx")
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    if "מגדל" in cell.value:
                        cell.value = "מגדל"
                    elif "קש חתמים" in cell.value:
                        cell.value = "קש חתמים"
    wb.save(fr"C:\ledugma\DEMO\{year}.xlsx")
    print(f"Updated {year}.xlsx")
