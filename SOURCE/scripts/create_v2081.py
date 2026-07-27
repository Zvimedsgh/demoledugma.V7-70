import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.080.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.081.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_081"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.081\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.081"\n')
    elif "Public Sub A00_SetupMainSheet_V2080()" in line:
        new_lines.append(line.replace("A00_SetupMainSheet_V2080", "A00_SetupMainSheet"))
    else:
        new_lines.append(line)

# Add FixExcel Sub at the end
fix_excel_code = """
' ==========================================
' FIX EXCEL ENVIRONMENT VARIABLES
' ==========================================
Public Sub FixExcel()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBoxU ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1488) & ChrW(1511) & ChrW(1505) & ChrW(1500) & " " & ChrW(1488) & ChrW(1493) & ChrW(1508) & ChrW(1505) & ChrW(1493) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & ".", vbInformation
End Sub
"""
new_lines.append(fix_excel_code)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.081 created with renamed Setup and FixExcel.")

