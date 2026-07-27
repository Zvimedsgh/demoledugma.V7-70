import win32com.client

try:
    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    
    with open("C:/LEVAV PROJECT/SOURCE/names_debug.txt", "w", encoding="utf-8") as f:
        f.write(f"Active Workbook: {wb.Name}\n")
        names_to_check = ["lst_period_type", "lst_half_year", "lst_quarter", "lst_month"]
        for name in names_to_check:
            try:
                n = wb.Names(name)
                f.write(f"{name} = {n.RefersTo}\n")
                
                # Try to get the values
                rn = excel.Range(name)
                f.write(f"Values for {name}:\n")
                if rn.Count == 1:
                    f.write(f"  {rn.Value}\n")
                else:
                    for cell in rn:
                        f.write(f"  {cell.Value}\n")
            except Exception as e:
                f.write(f"Error reading {name}: {str(e)}\n")
                
    print("Success")
except Exception as e:
    print(f"Failed: {str(e)}")
