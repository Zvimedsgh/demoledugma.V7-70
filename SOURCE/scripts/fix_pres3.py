import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.058.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.059.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_build_presentation = False
skip_lines = 0

for i, line in enumerate(lines):
    if skip_lines > 0:
        skip_lines -= 1
        continue
        
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_059"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.059\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.059"\n')
    elif "Public Sub BuildPresentation()" in line:
        in_build_presentation = True
        new_lines.append(line)
    elif in_build_presentation and "End Sub" in line:
        in_build_presentation = False
        new_lines.append(line)
    elif in_build_presentation and 'clientName = Trim$(CStr(wsMain.Range("rngClientName").Value2))' in line:
        new_lines.append('    If isDemoMode Then\n')
        new_lines.append('        clientName = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1491) & ChrW(1502) & ChrW(1493) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)\n')
        new_lines.append('    Else\n')
        new_lines.append('        clientName = Trim$(CStr(wsMain.Range("rngClientName").Value2))\n')
        new_lines.append('    End If\n')
    elif in_build_presentation and 'If Not bSaved Then' in line and 'askMsg =' in lines[i+2]:
        # This is the end-of-function block we want to replace
        new_lines.append('    If bSaved Then\n')
        new_lines.append('        askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1493) & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & vbCrLf & reportsFolder\n')
        new_lines.append('        MsgBoxU askMsg, vbOKOnly + vbInformation\n')
        new_lines.append('    Else\n')
        new_lines.append(lines[i+1])
        new_lines.append(lines[i+2])
        new_lines.append(lines[i+3])
        new_lines.append(lines[i+4])
        new_lines.append(lines[i+5])
        skip_lines = 5
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.059")

