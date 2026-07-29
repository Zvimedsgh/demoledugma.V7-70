import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.03.bas', 'r', encoding='utf-8') as f:
    content = f.read()

font_old = """' ---- Set RTL and font size 14 for ALL sheets in workbook ----
4720 For Each wsLoop In ThisWorkbook.Worksheets
4730 wsLoop.DisplayRightToLeft = True
4740 If wsLoop.Name <> CONTROL_SHEET_NAME() Then wsLoop.Cells.Font.Size = 14
4750 Next wsLoop"""

font_new = """' ---- Set RTL and font size 14 for ALL sheets in workbook ----
4720 For Each wsLoop In ThisWorkbook.Worksheets
4730 wsLoop.DisplayRightToLeft = True
     On Error Resume Next
4740 If wsLoop.Name <> CONTROL_SHEET_NAME() Then wsLoop.Cells.Font.Size = 14
     On Error GoTo ERR_HANDLER
4750 Next wsLoop"""

content = content.replace(font_old, font_new)

content = content.replace('APP_VERSION As String = "3.03"', 'APP_VERSION As String = "3.04"')
content = content.replace('VERSION: V3.03', 'VERSION: V3.04')
content = content.replace('Error in V3.03!', 'Error in V3.04!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_03"', 'Attribute VB_Name = "Goren_Claude_V3_04"')

changelog = """' CHANGES IN 3.04:
'   - BUGFIX: Added On Error Resume Next to Font.Size loop in A00 to prevent crash on protected sheets (Error 1004).
"""
content = content.replace("' CHANGES IN 3.03:", changelog + "' CHANGES IN 3.03:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.04.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.04')
