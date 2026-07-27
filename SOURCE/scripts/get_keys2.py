import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# Open the workbook
wb_path = r"C:\Users\Zvi\OneDrive - Zvi Goren א נ\ב ס\Demo_Reports_Syatem_V7.70  (2).xlsm"
try:
    wb = excel.Workbooks.Open(wb_path)
    
    ws = wb.Sheets("הגדרות_מיפוי")
    print("Found הגדרות_מיפוי! Extracting values:")
    for r in range(2, 50):
        he_name = ws.Cells(r, 1).Value
        col_let = ws.Cells(r, 2).Value
        key = ws.Cells(r, 4).Value
        if he_name and key:
            print(f"Row {r}: {he_name} -> {key} (Col {col_let})")
            
    wb.Close(False)
except Exception as e:
    print(f"Error: {e}")
finally:
    excel.Quit()
