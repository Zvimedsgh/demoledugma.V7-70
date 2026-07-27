import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.058.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.059.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []

for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_059"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.059\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.059"\n')
    elif 'clientName = Trim$(CStr(wsMain.Range("rngClientName").Value2))' in line:
        new_lines.append('    If isDemoMode Then\n')
        new_lines.append('        clientName = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1491) & ChrW(1502) & ChrW(1493) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)\n')
        new_lines.append('    Else\n')
        new_lines.append('        clientName = Trim$(CStr(wsMain.Range("rngClientName").Value2))\n')
        new_lines.append('    End If\n')
    elif 'Dim askMsg As String' in line and "BuildPresentation" in "".join(lines[max(0, i-50):i]):
        new_lines.append(line)
        # We need to replace the If Not bSaved block with a full If bSaved Then ... Else ...
        # But wait, it's easier to just inject the success message here.
    elif 'If Not bSaved Then' in line and "BuildPresentation" in "".join(lines[max(0, i-50):i]):
        new_lines.append('    If bSaved Then\n')
        new_lines.append('        askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1493) & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & vbCrLf & reportsFolder\n')
        new_lines.append('        MsgBoxU askMsg, vbOKOnly + vbInformation\n')
        new_lines.append('    Else\n')
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.059")

