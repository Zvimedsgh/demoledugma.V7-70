import win32com.client
try:
    xl = win32com.client.GetActiveObject("Excel.Application")
    wb = xl.ActiveWorkbook
    if wb is None:
        print("No active workbook")
    else:
        for ws in wb.Worksheets:
            if ws.Name == "NIHUL" or ws.Name == "הגדרות_פרמטרים" or "הגדרות" in ws.Name:
                print(f"Found sheet: {ws.Name}")
                print(f"ScrollArea: {ws.ScrollArea}")
                print(f"Shapes count: {ws.Shapes.Count}")
                for i in range(1, min(ws.Shapes.Count + 1, 10)):
                    shp = ws.Shapes(i)
                    print(f" Shape {i}: {shp.Name} (Type: {shp.Type}), w:{shp.Width}, h:{shp.Height}")
except Exception as e:
    print(e)
