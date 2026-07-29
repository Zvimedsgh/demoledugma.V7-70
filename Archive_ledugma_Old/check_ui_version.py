import os
import win32com.client as win32

try:
    excel = win32.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    ws = wb.Worksheets("דף הבית")
    
    val_a22 = ws.Range("A22").Value
    print(f"Cell A22 Value: {val_a22}")
    
    for shp in ws.Shapes:
        try:
            txt = shp.TextFrame2.TextRange.Text
            if '2.' in txt:
                print(f"Shape {shp.Name} contains: {txt}")
        except:
            pass
except Exception as e:
    print(f"Error: {e}")
