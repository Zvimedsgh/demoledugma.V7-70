import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.257.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.257"', 'APP_VERSION As String = "2.258"')
content = content.replace('VERSION: V2.257', 'VERSION: V2.258')
content = content.replace('Error in V2.257!', 'Error in V2.258!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_257"', 'Attribute VB_Name = "Goren_Claude_V2_258"')

# Update Changelog
changelog = """' CHANGES IN 2.258:
'   - UI: Moved WhatsApp credit to C16 and Version text to A16 to fit laptop screens without zooming out.
'   - UI: Restored zoom behavior to A1:L20.
'   - UX: Changed default parking cursor from F10 to G10 (the actual input cell).
"""
content = content.replace("' CHANGES IN 2.257:", changelog + "' CHANGES IN 2.257:")

# Fix the zoom back to normal
content = content.replace('wsMain.Range("A1:P22").Select', 'wsMain.Range("A1:L20").Select')

# Move version text to A16
content = content.replace('Range("A22")', 'Range("A16")')

# Move WhatsApp credit to C16:I16
content = content.replace('With wsMain.Range("D19:I19")', 'With wsMain.Range("C16:I16")')

# Change Goto F10 to Goto G10
content = content.replace('Application.Goto wsMain.Range("F10")', 'Application.Goto wsMain.Range("G10")')
content = content.replace('Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("F10")', 'Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G10")')
content = content.replace('Application.Goto wsMainUI.Range("F10")', 'Application.Goto wsMainUI.Range("G10")')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.258.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 258')
