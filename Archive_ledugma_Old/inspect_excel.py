import os
import win32com.client as win32

found_file = None
for root, dirs, files in os.walk(r'C:\Users\Zvi'):
    if 'OneDrive' in root:
        for f in files:
            if f.lower() == 'demo_reports_syatem_v7.70  (2).xlsm':
                found_file = os.path.join(root, f)
                break
    if found_file:
        break

if not found_file:
    print("File not found.")
else:
    print(f"Opening: {found_file}")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(found_file)
    
    print("\n--- Worksheets ---")
    for ws in wb.Worksheets:
        vis = "Visible" if ws.Visible == -1 else ("Hidden" if ws.Visible == 0 else "VeryHidden")
        print(f"Sheet: {ws.Name} | {vis}")
        
    ws_home = None
    try:
        ws_home = wb.Worksheets("דף הבית")
    except:
        pass
        
    if ws_home:
        print("\n--- Shapes on דף הבית ---")
        for shp in ws_home.Shapes:
            # shp.Visible is usually -1 for True, 0 for False in VBA
            vis = "Visible" if shp.Visible else "Hidden"
            text = ""
            if shp.Type in (1, 17): # msoAutoShape or msoTextBox
                try:
                    text = shp.TextFrame2.TextRange.Text[:30].replace('\r', ' ').replace('\n', ' ')
                except:
                    pass
            print(f"Shape: {shp.Name} | Type: {shp.Type} | {vis} | Left: {shp.Left}, Top: {shp.Top} | Text: {text}")
            
            # If it's hidden, let's unhide it just in case
            if not shp.Visible:
                print(f"  -> Unhiding {shp.Name}!")
                shp.Visible = -1
                shp.Left = 100
                shp.Top = 100
    
    wb.Save()
    wb.Close()
    excel.Quit()
