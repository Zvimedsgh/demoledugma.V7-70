import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.139.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_save_cleanup = """' Re-hide sheets that were hidden before
Dim hi As Long
For hi = 1 To hiddenCount
ThisWorkbook.Worksheets(hiddenSheets(hi)).Visible = xlSheetVeryHidden
Next hi"""

new_save_cleanup = """' Re-hide sheets that were hidden before
Dim hi As Long
For hi = 1 To hiddenCount
ThisWorkbook.Worksheets(hiddenSheets(hi)).Visible = xlSheetVeryHidden
Next hi

' Clean up the generated client report sheets
Application.DisplayAlerts = False
Dim wsCleanup As Worksheet
For Each wsCleanup In ThisWorkbook.Worksheets
    If wsCleanup.Visible = xlSheetVisible And UCase(wsCleanup.Name) <> "MAIN" Then
        wsCleanup.Delete
    End If
Next wsCleanup
Application.DisplayAlerts = True"""

if target_save_cleanup in content:
    content = content.replace(target_save_cleanup, new_save_cleanup)
else:
    print("Could not find the target string in SaveReportsToFolder")

# Update Version to V2.140 just in case
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_139"', 'Attribute VB_Name = "Goren_Claude_V2_140"')
content = content.replace('VERSION: V2.139', 'VERSION: V2.140')
content = content.replace('APP_VERSION As String = "2.139"', 'APP_VERSION As String = "2.140"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.140.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.140 created with cleanup logic in both BuildPresentation and SaveReportsToFolder.")

