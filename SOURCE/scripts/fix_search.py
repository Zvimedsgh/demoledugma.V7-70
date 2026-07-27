import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """    ' If only 1 result, select it directly
    If dict.Count = 1 Then
        Dim wsMain As Worksheet
        Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Range("rngClientName").Value = dict.keys()(0)
        ' Clean up search sheet
        Application.DisplayAlerts = False
        wsSearch.Delete
        Application.DisplayAlerts = True
        wsMain.Activate
        MsgBoxU ChrW(1504) & ChrW(1489) & ChrW(1495) & ChrW(1512) & ": " & dict.keys()(0), vbInformation  ' "????: [name]"
        Exit Sub
    End If"""

new_str = """    ' Previously auto-selected if dict.Count = 1, but removed per user request"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed Client Search")
else:
    print("Could not find Client Search string")
