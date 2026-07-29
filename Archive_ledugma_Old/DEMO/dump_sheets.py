import openpyxl

wb = openpyxl.load_workbook(r'C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm', data_only=True)
with open('sheets.txt', 'w', encoding='utf-8') as f:
    for sn in wb.sheetnames:
        f.write(f"{sn}\n")
