import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.156.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub HideWorkSheets" in line:
        insert_idx = i
        break

new_macro = """
Public Sub ShowResultSheets()
    ' Shows only the result sheets (companies, branches, agents, etc.)
    Dim ws As Worksheet
    Dim resNames As String
    resNames = "|" & SHEET_COMPANIES() & "|" & SHEET_AGENTS() & "|" & SHEET_BRANCH() & "|" & SHEET_MAINBRANCH() & "|" & SHEET_TELLERS() & "|" & SHEET_MONTHS() & "|" & SHEET_SUMMARY() & "|"
    
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Visible = xlSheetVisible
    
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, resNames, "|" & ws.Name & "|", vbTextCompare) > 0 Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
    
    ' Go to the first result sheet if it exists
    If SheetExists(SHEET_COMPANIES()) Then
        ThisWorkbook.Worksheets(SHEET_COMPANIES()).Activate
    End If
End Sub
"""

lines.insert(insert_idx, new_macro)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("ShowResultSheets macro added.")
