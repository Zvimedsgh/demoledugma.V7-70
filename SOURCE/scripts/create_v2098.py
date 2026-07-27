import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.097.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.098.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_097"', 'Attribute VB_Name = "Goren_Claude_V2_098"')
content = content.replace("' VERSION: V2.097", "' VERSION: V2.098")
content = content.replace('Private Const APP_VERSION As String = "2.097"', 'Private Const APP_VERSION As String = "2.098"')

target = """570     lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
580     If lastRow < 2 Then Err.Raise vbObjectError + 1006, "BuildReview", "NO DATA ROWS IN SOURCE"

' Header validation (minimum 25 cols)
590     Dim maxCol As Long
600     maxCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column"""

replacement = """570     lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
580     If lastRow < 2 Then Err.Raise vbObjectError + 1006, "BuildReview", "NO DATA ROWS IN SOURCE"

' Header validation (minimum 25 cols)
590     Dim maxCol As Long
600     maxCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
        If maxCol < 1 Then maxCol = 100
        srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, maxCol)).Value2
"""

content = content.replace(target, replacement)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.098 created.")
