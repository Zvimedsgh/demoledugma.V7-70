import win32com.client

try:
    excel = win32com.client.GetActiveObject("Excel.Application")
    for wb in excel.Workbooks:
        print(f"Found workbook: {wb.Name}")
        try:
            ws = wb.Sheets("הגדרות_מיפוי")
            print("Found הגדרות_מיפוי! Extracting values:")
            for r in range(2, 50):
                he_name = ws.Cells(r, 1).Value
                col_let = ws.Cells(r, 2).Value
                key = ws.Cells(r, 4).Value
                if he_name and key:
                    print(f"Row {r}: {he_name} -> {key} (Col {col_let})")
        except Exception as e:
            pass
except Exception as e:
    print(f"Error accessing Excel: {e}")
