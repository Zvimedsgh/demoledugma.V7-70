def create_v957():
    with open(r'C:\ledugma\DEMO\V9.56_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # Find the FormatConditions block and replace it entirely
    old_block = """wsMain.Range("G7").FormatConditions.Delete
wsMain.Range("G7").FormatConditions.Add Type:=xlExpression, _
Formula1:="=AND(LEN($G$5)>0,$G$5<>"" & mgmtRef & ""!$Z$1,$G$5<>"" & mgmtRef & ""!$Z$2,OR($G$6="""""""",$G$6="" & mgmtRef & ""!$Z$2))"
wsMain.Range("G7").FormatConditions(wsMain.Range("G7").FormatConditions.count).Interior.Color = RGB(255, 215, 0)
wsMain.Range("G10").FormatConditions.Delete
wsMain.Range("G11").FormatConditions.Add Type:=xlExpression, _
Formula1:="=AND(LEN($G$9)>0,$G$9<>"" & mgmtRef & ""!$Z$2,OR($G$10="""""""",$G$10="" & mgmtRef & ""!$Z$2))"
wsMain.Range("G10").FormatConditions(1).Interior.Color = RGB(255, 215, 0)"""

    # We will use string manipulation to be safe
    # It starts at: wsMain.Range("G7").FormatConditions.Delete
    # It ends before: On Error GoTo ERR_HANDLER
    start_idx = code.find('wsMain.Range("G7").FormatConditions.Delete')
    end_idx = code.find('On Error GoTo ERR_HANDLER', start_idx)

    new_block = """wsMain.Range("G7").FormatConditions.Delete
        wsMain.Range("G7").FormatConditions.Add Type:=xlExpression, Formula1:="=($G$7="""")*($G$6<>"""")"
        wsMain.Range("G7").FormatConditions(1).Interior.Color = RGB(255, 215, 0)
        
        wsMain.Range("G10").FormatConditions.Delete
        wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, Formula1:="=($G$10="""")*($G$9<>"""")"
        wsMain.Range("G10").FormatConditions(1).Interior.Color = RGB(255, 215, 0)
        """
        
    code = code[:start_idx] + new_block + code[end_idx:]

    with open(r'C:\ledugma\DEMO\V9.57_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v957()
