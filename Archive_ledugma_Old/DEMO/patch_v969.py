import datetime

def create_v969():
    with open(r'C:\ledugma\DEMO\V9.68_CopyPaste.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    version_string = f"9.69 ({date_str} {time_str})"

    # Update the header block
    code = code.replace("' VERSION: 9.68", f"' VERSION: {version_string}")
    code = code.replace("'          VERSION 9.68", f"'          VERSION 9.69")
    code = code.replace("' CHANGES IN 9.68:", f"' CHANGES IN 9.69:\n' - Added EnsureNamedRanges to ClearHomePageSelection to automatically fix old workbooks without running SetupMainSheet\n' CHANGES IN 9.68:")
    code = code.replace("V9.68", "V9.69")
    code = code.replace('Public Const SYSTEM_VERSION = "9.68"', 'Public Const SYSTEM_VERSION = "9.69"')

    ensure_ranges = """Private Sub EnsureNamedRanges()
    Dim wsMgmt As Worksheet
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    
    ' Populate half year
    wsMgmt.Cells(2, 18).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492)
    wsMgmt.Cells(3, 18).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1492)
    
    ' Populate quarters
    wsMgmt.Cells(5, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503)
    wsMgmt.Cells(6, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(7, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497)
    wsMgmt.Cells(8, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)
    
    ' Populate months
    wsMgmt.Cells(10, 18).Value = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    wsMgmt.Cells(11, 18).Value = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    wsMgmt.Cells(12, 18).Value = ChrW(1502) & ChrW(1512) & ChrW(1509)
    wsMgmt.Cells(13, 18).Value = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
    wsMgmt.Cells(14, 18).Value = ChrW(1502) & ChrW(1488) & ChrW(1497)
    wsMgmt.Cells(15, 18).Value = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(16, 18).Value = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
    wsMgmt.Cells(17, 18).Value = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
    wsMgmt.Cells(18, 18).Value = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(19, 18).Value = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(20, 18).Value = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(21, 18).Value = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    
    ' Define Named Ranges
    On Error Resume Next
    ThisWorkbook.names("lst_half_year").Delete
    ThisWorkbook.names("lst_quarter").Delete
    ThisWorkbook.names("lst_month").Delete
    On Error GoTo 0
    ThisWorkbook.names.Add Name:="lst_half_year", RefersTo:="='" & wsMgmt.Name & "'!$R$2:$R$3"
    ThisWorkbook.names.Add Name:="lst_quarter", RefersTo:="='" & wsMgmt.Name & "'!$R$5:$R$8"
    ThisWorkbook.names.Add Name:="lst_month", RefersTo:="='" & wsMgmt.Name & "'!$R$10:$R$21"
End Sub

Public Sub ClearHomePageSelection()
    EnsureNamedRanges
"""

    code = code.replace("Public Sub ClearHomePageSelection()\n", ensure_ranges)

    with open(r'C:\ledugma\DEMO\V9.69_Final.txt', 'w', encoding='windows-1255') as f:
        f.write(code)

if __name__ == '__main__':
    create_v969()
