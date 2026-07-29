import datetime

def create_v970():
    with open(r'C:\ledugma\DEMO\V9.69_Final.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.70 ({date_str} {time_str})"

    # Update version
    code = code.replace("9.69", "9.70")
    code = code.replace("V9.69", "V9.70")

    # Fix Unprotect
    unprotect_old = """        Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Unprotect SHEET_PROTECT_PWD"""
    
    unprotect_new = """        Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        On Error Resume Next
        wsMain.Unprotect SHEET_PROTECT_PWD
        wsMain.Unprotect ""
        On Error GoTo ERR_HANDLER"""
    
    code = code.replace(unprotect_old, unprotect_new)

    # Fix Else block in UpdatePeriodDropdown
    else_old = """        Else
            GoTo CLEAN_EXIT
        End If"""
    
    else_new = """        Else
            MsgBox ChrW(1492) & ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1489) & "-" & "G6" & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1494) & ChrW(1493) & ChrW(1492) & ChrW(1492) & ":" & " " & periodType, vbExclamation
            GoTo CLEAN_EXIT
        End If"""
    
    code = code.replace(else_old, else_new)

    # Fix ERR_HANDLER
    err_old = """ERR_HANDLER:

End Sub"""
    
    err_new = """ERR_HANDLER:
        MsgBox "Error in G7: " & Err.Description, vbCritical
End Sub"""
    
    code = code.replace(err_old, err_new)

    with open(r'C:\ledugma\DEMO\V9.70_Final.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v970()
