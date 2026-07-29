import openpyxl

filepath = r"C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm"
wb = openpyxl.load_workbook(filepath, data_only=True)
ws = wb["ראשי"]

for row in range(1, 25):
    f_val = ws.cell(row=row, column=6).value
    g_val = ws.cell(row=row, column=7).value
    if f_val or g_val:
        print(f"Row {row}: F={f_val}, G={g_val}")

print("Shapes in 'ראשי':")
# openpyxl doesn't fully support reading shape coordinates easily, but we can see drawing elements
if ws._images:
    for img in ws._images:
        print(f"Image anchor: {img.anchor}")
