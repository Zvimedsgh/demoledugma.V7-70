import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.100.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

def repl_hide(match):
    return """Public Sub HideWorkSheets()
' Hides ALL sheets except home sheet (daf habait)
Dim ws As Worksheet
Dim ctrlName As String

ctrlName = CONTROL_SHEET_NAME()

' Make sure home sheet is visible before hiding others
ThisWorkbook.Worksheets(ctrlName).Visible = xlSheetVisible

For Each ws In ThisWorkbook.Worksheets
If UCase$(ws.Name) <> UCase$(ctrlName) Then
    ws.Visible = xlSheetVeryHidden
End If
Next ws

' Activate home sheet
ThisWorkbook.Worksheets(ctrlName).Activate
End Sub"""

content = re.sub(r'Public Sub HideWorkSheets\(\).*?End Sub', repl_hide, content, flags=re.DOTALL)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.100 updated.")
