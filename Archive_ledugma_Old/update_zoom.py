import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.276.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for line in lines:
    if 'Public Sub ToggleDisplayMode()' in line:
        skip = True
        new_lines.append(line)
        new_lines.append('    On Error Resume Next\n')
        new_lines.append('    Dim wsMain As Worksheet\n')
        new_lines.append('    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1508) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))\n')
        new_lines.append('    If wsMain Is Nothing Then Exit Sub\n')
        new_lines.append('    wsMain.Unprotect Password:="1234"\n')
        new_lines.append('    Dim shpToggleLayout As Shape\n')
        new_lines.append('    Set shpToggleLayout = wsMain.Shapes("shpToggleLayout")\n')
        new_lines.append('    If ActiveWindow.Zoom > 80 Then\n')
        new_lines.append('        ActiveWindow.Zoom = 75\n')
        new_lines.append('        If Not shpToggleLayout Is Nothing Then\n')
        new_lines.append('            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1512) & ChrW(1495) & ChrW(1489)\n')
        new_lines.append('        End If\n')
        new_lines.append('    Else\n')
        new_lines.append('        ActiveWindow.Zoom = 100\n')
        new_lines.append('        If Not shpToggleLayout Is Nothing Then\n')
        new_lines.append('            shpToggleLayout.TextFrame2.TextRange.Text = ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1496) & ChrW(1493) & ChrW(1508)\n')
        new_lines.append('        End If\n')
        new_lines.append('    End If\n')
        new_lines.append('    wsMain.Protect Password:="1234"\n')
        new_lines.append('End Sub\n')
        continue
    
    if skip:
        if 'End Sub' in line:
            skip = False
        continue
    
    new_lines.append(line)

content = "".join(new_lines)
content = content.replace('APP_VERSION As String = "2.276"', 'APP_VERSION As String = "2.277"')
content = content.replace('VERSION: V2.276', 'VERSION: V2.277')
content = content.replace('Error in V2.276!', 'Error in V2.277!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_276"', 'Attribute VB_Name = "Goren_Claude_V2_277"')

changelog = """' CHANGES IN 2.277:
'   - UI: Redesigned ToggleDisplayMode to purely adjust ActiveWindow.Zoom between 100% and 75%, eliminating complex DOM repositioning logic.
"""
content = content.replace("' CHANGES IN 2.276:", changelog + "' CHANGES IN 2.276:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.277.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 277')
