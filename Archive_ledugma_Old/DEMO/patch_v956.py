import re

def create_v956():
    with open(r'C:\ledugma\DEMO\V9.55_CopyPaste.txt', 'r', encoding='windows-1255') as f:
        code = f.read()

    # 1. Fix column G width override
    code = code.replace(
        'wsMain.Columns("F:G").ColumnWidth = 16   \' Parameter Table (reverted to F:G)',
        'wsMain.Columns("F").ColumnWidth = 16\n    wsMain.Columns("G").ColumnWidth = 25'
    )

    # 2. Fix Conditional Formatting formulas to avoid language (AND) bugs
    # For G7
    code = code.replace(
        'Formula1:="=AND($G$7="""", $G$6<>"""")"',
        'Formula1:="=($G$7="""")*($G$6<>"""")"'
    )
    # For G10
    code = code.replace(
        'Formula1:="=AND($G$10="""", $G$9<>"""")"',
        'Formula1:="=($G$10="""")*($G$9<>"""")"'
    )

    # 3. Update Period/Filter Dropdowns to NOT rely on named ranges just in case that's failing
    # In UpdatePeriodDropdown:
    code = code.replace(
        'periodType = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))',
        'periodType = Trim$(CStr(wsMain.Range("G6").Value2))'
    )
    code = code.replace(
        'wsMain.Range("rngPeriodValue").Value = ""',
        'wsMain.Range("G7").Value = ""'
    )
    code = code.replace(
        'wsMain.Range("rngPeriodValue").Validation.Delete',
        'wsMain.Range("G7").Validation.Delete'
    )
    code = code.replace(
        'wsMain.Range("rngPeriodValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName',
        'wsMain.Range("G7").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName'
    )
    
    # In UpdateFilterValueDropdown:
    code = code.replace(
        'filterType = Trim$(CStr(wsMain.Range("rngFilterType").Value2))',
        'filterType = Trim$(CStr(wsMain.Range("G9").Value2))'
    )
    code = code.replace(
        'wsMain.Range("rngFilterValue").Value = ""',
        'wsMain.Range("G10").Value = ""'
    )
    code = code.replace(
        'wsMain.Range("rngFilterValue").Validation.Delete',
        'wsMain.Range("G10").Validation.Delete'
    )
    code = code.replace(
        'wsMain.Range("rngFilterValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName',
        'wsMain.Range("G10").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName'
    )

    with open(r'C:\ledugma\DEMO\V9.56_CopyPaste.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v956()
