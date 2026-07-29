import os
import win32com.client as win32

file_path = None
for root, dirs, files in os.walk(r'C:\Users\Zvi'):
    if 'OneDrive' in root:
        for f in files:
            if f == 'Demo_Reports_Syatem_V7.70  (2)2.xlsm':
                file_path = os.path.join(root, f)
                break
    if file_path:
        break

if file_path:
    print(f"Found: {file_path}")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(file_path, ReadOnly=True)
    print("Sheets and their visibility:")
    for ws in wb.Worksheets:
        print(f"  {ws.Name}: Visible={ws.Visible}")
    print(f"Active Sheet: {wb.ActiveSheet.Name}")
    wb.Close(False)
    excel.Quit()
else:
    print("File not found locally.")
