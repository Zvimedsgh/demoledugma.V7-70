import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.140.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(r"(' Re-hide sheets that were hidden before\s*Dim hi As Long\s*For hi = 1 To hiddenCount\s*ThisWorkbook\.Worksheets\(hiddenSheets\(hi\)\)\.Visible = xlSheetVeryHidden\s*Next hi)")

new_save_cleanup = """\g<1>

' Clean up the generated client report sheets
Application.DisplayAlerts = False
Dim wsCleanup As Worksheet
For Each wsCleanup In ThisWorkbook.Worksheets
    If wsCleanup.Visible = xlSheetVisible And UCase(wsCleanup.Name) <> "MAIN" And UCase(wsCleanup.Name) <> UCase(CONTROL_SHEET_NAME()) Then
        wsCleanup.Delete
    End If
Next wsCleanup
Application.DisplayAlerts = True"""

if pattern.search(content):
    content = pattern.sub(new_save_cleanup, content)
    print("Found and replaced target block!")
else:
    print("Could not find block with regex.")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

