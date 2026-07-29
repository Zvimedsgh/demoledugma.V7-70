import openpyxl

filepath = r"C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm"
wb = openpyxl.load_workbook(filepath, data_only=True)
ws = wb.worksheets[0]

for row in range(1, 25):
    for col in range(1, 10):
        val = ws.cell(row=row, column=col).value
        if val:
            print(f"Row {row}, Col {col}: {val}")

print("Shapes in 'ראשי':")
# openpyxl doesn't fully support reading shape coordinates easily, but we can see drawing elements
if ws._images:
    for img in ws._images:
        print(f"Image anchor: {img.anchor}")
