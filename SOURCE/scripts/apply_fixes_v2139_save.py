import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.139.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_save_exit = """' Restore hidden sheets in the original workbook
For i = 1 To hiddenCount
On Error Resume Next
ThisWorkbook.Worksheets(hiddenSheets(i)).Visible = xlSheetVeryHidden
Next i

Application.EnableEvents = True"""

save_cleanup = """' Restore hidden sheets in the original workbook
For i = 1 To hiddenCount
On Error Resume Next
ThisWorkbook.Worksheets(hiddenSheets(i)).Visible = xlSheetVeryHidden
Next i

' Clean up the exported report sheets from the original workbook
Application.DisplayAlerts = False
Dim wsCleanup As Worksheet
For Each wsCleanup In ThisWorkbook.Worksheets
    If wsCleanup.Visible = xlSheetVisible And wsCleanup.Name <> "Main" Then
        wsCleanup.Delete
    End If
Next wsCleanup
Application.DisplayAlerts = True

Application.EnableEvents = True"""

if target_save_exit in content:
    content = content.replace(target_save_exit, save_cleanup)
else:
    print("Could not find SaveReportsToFolder exit block!")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.139.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.139 updated with SaveReportsToFolder cleanup.")

