import os
import glob
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

base_dir = r"C:\Users\Zvi"
found = False

for root, dirs, files in os.walk(base_dir):
    for f in files:
        if "Demo_Reports_Syatem_V7.70" in f and f.endswith(".xlsm"):
            wb_path = os.path.join(root, f)
            print(f"Found: {wb_path}")
            try:
                wb = excel.Workbooks.Open(wb_path)
                ws = wb.Sheets("הגדרות_מיפוי")
                for r in range(2, 50):
                    he_name = ws.Cells(r, 1).Value
                    col_let = ws.Cells(r, 2).Value
                    key = ws.Cells(r, 4).Value
                    if he_name and key:
                        print(f"Row {r}: {he_name} -> {key} (Col {col_let})")
                wb.Close(False)
                found = True
                break
            except Exception as e:
                print(f"Error reading {wb_path}: {e}")

excel.Quit()
