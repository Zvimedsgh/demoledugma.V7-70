import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.079.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.080.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_080"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.080\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.080"\n')
    elif "Public Sub A00_SetupMainSheet_V2079()" in line:
        new_lines.append(line.replace("A00_SetupMainSheet_V2079", "A00_SetupMainSheet_V2080"))
    elif "1110 ThisWorkbook.Names.Add \"map_filter\", wsTemp.Range(\"G1:H5\")" in line:
        new_lines.append(line)
        new_lines.append('        1111 Dim yIdx As Long\n')
        new_lines.append('        1112 For yIdx = 1 To 20\n')
        new_lines.append('        1113     wsTemp.Cells(yIdx, 9).Value = 2020 + yIdx\n')
        new_lines.append('        1114 Next yIdx\n')
        new_lines.append('        1115 ThisWorkbook.Names.Add "lst_years", wsTemp.Range("I1:I20")\n')
    elif 'Operator:=xlBetween, Formula1:="=lst_years"' in line:
        new_lines.append('        1775 On Error Resume Next\n')
        new_lines.append(line.replace("Operator:=xlBetween, ", ""))
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.080 created with lst_years defined.")

