import win32com.client

try:
    xl = win32com.client.GetActiveObject("Excel.Application")
    wb = xl.ActiveWorkbook
    if wb is None:
        print("No active workbook found.")
    else:
        print("Workbook:", wb.Name)
        ws = None
        for sheet in wb.Worksheets:
            if sheet.Name == "NIHUL":
                ws = sheet
                break
        
        if ws is None:
            print("NIHUL sheet not found!")
        else:
            ws.Unprotect("Z961814r")
            
            # Find the first empty row in column C (Param Name)
            last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row # xlUp
            
            # Write AGENCY_NAME
            ws.Cells(last_row + 1, 3).Value = "AGENCY_NAME"
            ws.Cells(last_row + 1, 4).Value = "סוכנות ביטוח לבב"
            
            # Write DEMO_AGENCY_NAME
            ws.Cells(last_row + 2, 3).Value = "DEMO_AGENCY_NAME"
            ws.Cells(last_row + 2, 4).Value = "לדוגמא סוכנות לביטוח"
            
            print("Successfully added parameters to NIHUL sheet!")
            
except Exception as e:
    print("Error:", str(e))
