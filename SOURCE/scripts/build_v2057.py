import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        lines[i] = 'Attribute VB_Name = "Goren_Claude_V2_057"\n'
    elif "VERSION: " in line:
        lines[i] = "' VERSION: V2.057\n"
    elif "Private Const APP_VERSION As String =" in line:
        lines[i] = 'Private Const APP_VERSION As String = "2.057"\n'
    
    # 1. Add MsgBox at top of BuildReview
    if "Public Sub BuildReview()" in line:
        lines[i] = line + '    MsgBoxU ChrW(1502) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " 2.057!!!", vbInformation\n'
        
    # 2. Relax IsProtectedSheet to catch anything with "DATA" (case-insensitive)
    if 'If UCase$(Left$(sName, 5)) = "DATA_" Then' in line:
        lines[i] = '    If InStr(1, UCase$(sName), "DATA") > 0 Then IsProtectedSheet = True: Exit Function\n'

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.057.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.057")
