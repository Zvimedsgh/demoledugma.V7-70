import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.095.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_095"', 'Attribute VB_Name = "Goren_Claude_V2_099"')
content = content.replace("' VERSION: V2.095", "' VERSION: V2.099")
content = content.replace('Private Const APP_VERSION As String = "2.095"', 'Private Const APP_VERSION As String = "2.099"')

target = """590     Dim maxCol As Long
600     maxCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column"""

replacement = """590     Dim maxCol As Long
600     maxCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If maxCol < 1 Then maxCol = 100
    srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, maxCol)).Value2"""

if target in content:
    content = content.replace(target, replacement)
    print("Replacement successful.")
else:
    print("Target not found. Let's look for parts of it.")

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.099 created.")
