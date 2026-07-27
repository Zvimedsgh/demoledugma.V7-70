import sys
import win32com.client

try:
    xl = win32com.client.GetObject(None, "Excel.Application")
    wb = xl.ActiveWorkbook
    ws = wb.Worksheets("דף הבית")
    print("Found sheet: דף הבית")
    
    shapes_info = []
    for shp in ws.Shapes:
        try:
            txt = shp.TextFrame.Characters().Text
            shapes_info.append({"name": shp.Name, "text": txt.strip(), "top": shp.Top, "left": shp.Left, "height": shp.Height})
        except:
            pass
            
    # Sort by Top to see order
    shapes_info.sort(key=lambda x: x["top"])
    for s in shapes_info:
        if s["text"]:
            print(f"Name: {s['name']}, Text: {s['text']}, Top: {s['top']}, Height: {s['height']}")
except Exception as e:
    print("Error:", str(e))
