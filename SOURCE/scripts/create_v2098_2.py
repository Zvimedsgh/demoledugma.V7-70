import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.097.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.098.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_097"', 'Attribute VB_Name = "Goren_Claude_V2_098"')
content = content.replace("' VERSION: V2.097", "' VERSION: V2.098")
content = content.replace('Private Const APP_VERSION As String = "2.097"', 'Private Const APP_VERSION As String = "2.098"')

def insert_srcdata(match):
    original = match.group(0)
    insertion = """
    ' Populate srcData array for ultra-fast processing
    Dim maxCol As Long
    maxCol = GetLastCol(wsSrc, 1)
    If maxCol < 1 Then maxCol = 100
    srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, maxCol)).Value2
"""
    return original + insertion

pattern = re.compile(r'(570\s+lastRow = wsSrc\.Cells\(wsSrc\.Rows\.Count, 1\)\.End\(xlUp\)\.Row\s*\n580\s+If lastRow < 2 Then Err\.Raise vbObjectError \+ 1006, "BuildReview", "NO DATA ROWS IN SOURCE"\s*\n)')
content = pattern.sub(insert_srcdata, content, count=1)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.098 created using regex.")
