with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

new_hide = """Sub HideWorkSheets()
' Hides ALL sheets except home sheet (daf habait)
Dim ws As Worksheet
Dim ctrlName As String

On Error Resume Next
ThisWorkbook.Unprotect "Z961814r"

ctrlName = CONTROL_SHEET_NAME()

' Make sure home sheet is visible before hiding others
ThisWorkbook.Worksheets(ctrlName).Visible = xlSheetVisible

For Each ws In ThisWorkbook.Worksheets
If UCase$(ws.Name) <> UCase$(ctrlName) Then
Dim hideIt As Boolean
hideIt = True

Dim wsName1 As String, wsName2 As String
wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
wsName2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)

If ws.Name = wsName1 Or ws.Name = wsName2 Or ws.Name = "Sheet2" Then
If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then
hideIt = False
ThisWorkbook.Unprotect "Z961814r"
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
ws.Move After:=ThisWorkbook.Worksheets(ctrlName)
End If
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' _
If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then
hideIt = False
ThisWorkbook.Unprotect "Z961814r"
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
End If
End If

If hideIt Then 
    ThisWorkbook.Unprotect "Z961814r"
    ws.Visible = xlSheetVeryHidden
End If
End If
Next ws

' Auto-update version string on the Home Page
ThisWorkbook.Worksheets(ctrlName).Range("A22").Value = ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " v2.222"

' Activate home sheet
ThisWorkbook.Worksheets(ctrlName).Activate
' --- ADMIN MODE: LOCK SYSTEM ---
ThisWorkbook.Worksheets(ctrlName).Protect Password:="Z961814r", UserInterfaceOnly:=True
ThisWorkbook.Protect Password:="Z961814r"
On Error GoTo 0
End Sub"""

# Replace the old HideWorkSheets with new_hide
text = re.sub(r'Sub HideWorkSheets\(\).*?End Sub', new_hide, text, flags=re.DOTALL | re.IGNORECASE)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Replaced HideWorkSheets successfully!")
