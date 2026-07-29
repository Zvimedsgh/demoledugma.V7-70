import win32com.client as win32
import time
import sys

success = False
for i in range(5):
    try:
        excel = win32.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook
        
        if wb is None:
            time.sleep(1)
            continue
            
        ws = None
        for sheet in wb.Worksheets:
            if sheet.Name == "דף הבית":
                ws = sheet
                break
                
        if ws is None:
            print("Home page not found.")
            sys.exit()
            
        found = False
        for shp in ws.Shapes:
            if shp.Name == "shpInstallMsg":
                print(f"Shape found!")
                print(f"Left: {shp.Left}")
                print(f"Top: {shp.Top}")
                print(f"Width: {shp.Width}")
                print(f"Height: {shp.Height}")
                # Print RGB info
                print(f"Fill RGB: {shp.Fill.ForeColor.RGB}")
                print(f"Font Color: {shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB}")
                found = True
                success = True
                break
                
        if not found:
            print("Shape 'shpInstallMsg' not found on the Home Page.")
            sys.exit()
            
        if success:
            break
            
    except Exception as e:
        print(f"Attempt {i+1} Error: {e}")
        time.sleep(1)

if not success:
    print("Could not retrieve shape properties via COM.")
