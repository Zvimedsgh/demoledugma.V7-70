import win32com.client
import os

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    
    ws.Range("A1").Value = "Header"
    ws.Range("A2").Value = "Client 1"
    ws.Range("A3").Value = "Client 2"
    
    rng = ws.Range("A2:A3")
    
    # Try the syntax from the macro
    wb.Names.Add("lst_clients", rng)
    
    # Check if the name exists and what it refers to
    name_obj = wb.Names("lst_clients")
    print(f"Name refers to: {name_obj.RefersTo}")
    
    wb.Close(False)
    excel.Quit()
    print("Success")
except Exception as e:
    print(f"Error: {e}")
    try:
        excel.Quit()
    except:
        pass
