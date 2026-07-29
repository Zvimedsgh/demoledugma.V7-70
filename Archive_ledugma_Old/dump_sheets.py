import win32com.client
with open('sheets.txt', 'w', encoding='utf-8') as f:
    try:
        xl = win32com.client.GetActiveObject('Excel.Application')
        for wb in xl.Workbooks:
            f.write(f'Workbook: {wb.Name}\n')
            for ws in wb.Worksheets:
                f.write(f'  Sheet: {ws.Name} (Visible: {ws.Visible})\n')
    except Exception as e:
        f.write(str(e))
