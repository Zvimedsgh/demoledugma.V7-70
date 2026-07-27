import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.100.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_099"', 'Attribute VB_Name = "Goren_Claude_V2_100"')
content = content.replace("' VERSION: V2.099", "' VERSION: V2.100")
content = content.replace('Private Const APP_VERSION As String = "2.099"', 'Private Const APP_VERSION As String = "2.100"')

# 1. Fix HideWorkSheets to hide everything except home page
old_hide = """Public Sub HideWorkSheets()
' Hides ALL sheets except home sheet (daf habait)
Dim ws As Worksheet
Dim ctrlName As String

ctrlName = CONTROL_SHEET_NAME()

' Make sure home sheet is visible before hiding others
ThisWorkbook.Worksheets(ctrlName).Visible = xlSheetVisible

For Each ws In ThisWorkbook.Worksheets
Select Case ws.Name
Case CONTROL_SHEET_NAME(), MANAGEMENT_SHEET_NAME(), SHEET_COMPANIES(), _
SHEET_AGENTS(), SHEET_BRANCH(), SHEET_MAINBRANCH(), _
SHEET_TELLERS(), SHEET_MONTHS(), SHEET_SUMMARY()

ws.Visible = xlSheetVisible
Case Else
ws.Visible = xlSheetVeryHidden
End Select
Next ws

' Activate home sheet
ThisWorkbook.Worksheets(ctrlName).Activate

End Sub"""

new_hide = """Public Sub HideWorkSheets()
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
content = content.replace(old_hide, new_hide)

# 2. Fix NavSettings_Menu to use SettingsGoToSheet
old_nav_menu = """Public Sub NavSettings_Menu()
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If wsMgmt Is Nothing Then Exit Sub
    
60  wsMgmt.Activate
70  Application.GoTo Reference:=wsMgmt.Range("A1"), Scroll:=True
End Sub"""

new_nav_menu = """Public Sub NavSettings_Menu()
    SettingsGoToSheet MANAGEMENT_SHEET_NAME()
End Sub"""
content = content.replace(old_nav_menu, new_nav_menu)

# 3. Fix NavSettings_Home to call HideWorkSheets
old_nav_home = """Public Sub NavSettings_Home()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Set wsMain = ThisWorkbook.Worksheets("Main")
    wsMain.Activate
    Application.Goto wsMain.Range("A1")
End Sub"""

new_nav_home = """Public Sub NavSettings_Home()
    Call HideWorkSheets
End Sub"""
content = content.replace(old_nav_home, new_nav_home)

# 4. Fix NavToIndex to call HideWorkSheets
old_nav_index = """Public Sub NavToIndex()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    ws.Visible = xlSheetVisible
    ws.Activate
    Application.Goto ws.Range("A1")
End Sub"""

new_nav_index = """Public Sub NavToIndex()
    Call HideWorkSheets
End Sub"""
content = content.replace(old_nav_index, new_nav_index)


# 5. Fix SendForReview to use ScreenUpdating and call HideWorkSheets
# We will do this via regex since it's a large sub.
def fix_send_start(match):
    return match.group(0) + "\n    Application.ScreenUpdating = False\n"
content = re.sub(r'(Public Sub SendForReview\(\)\s*\n10\s*On Error GoTo ERR_HANDLER\s*\n)', fix_send_start, content)

def fix_send_end(match):
    original = match.group(0)
    return """
    Call HideWorkSheets
    Application.ScreenUpdating = True
    """ + original
content = re.sub(r'(\s*950\s*Dim askSent As Long)', fix_send_end, content)

def fix_send_err(match):
    return match.group(0) + "\n    Application.ScreenUpdating = True\n"
content = re.sub(r'(ERR_HANDLER:\s*\nApplication\.EnableEvents = True\s*\nApplication\.DisplayAlerts = True\s*\n)', fix_send_err, content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.100 created.")
