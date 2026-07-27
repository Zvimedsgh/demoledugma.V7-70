import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.074.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.075.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_setup = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_075"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.075\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.075"\n')
    elif "1680 With wsMain.Range(\"G3:G4\").Validation" in line:
        new_lines.append('        1680 On Error Resume Next\n')
        new_lines.append('        1681 With wsMain.Range("G3:G4").Validation\n')
    elif "1690 .Delete" in line:
        new_lines.append(line)
    elif "1700 .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:=\"=FALSE\"" in line:
        new_lines.append('        1700 .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=0"\n')
    elif "1710 .ErrorMessage =" in line:
        new_lines.append('        1705 .ErrorTitle = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492)\n')
        new_lines.append(line)
    elif "1720 .ShowError = True" in line:
        new_lines.append(line)
    elif "1730 End With" in line:
        new_lines.append(line)
        new_lines.append('        1731 On Error GoTo ERR_HANDLER\n')
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.075 created with safe validation block.")

