import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.174.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.175.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if line.strip() == 'On Error Resume Next' and lines[i+1].strip() == 'wsMain.Shapes("shpDemoMsgText").Delete' and lines[i+3].strip() == 'Dim shpDemoMsg As Shape':
        skip = True
        new_demo_ui = """    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo 0
    With wsMain.Range("E19:K20")
        .Merge
        .Value = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ": 054-6677396"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 255)
    End With\n"""
        new_lines.append(new_demo_ui)
        continue

    if skip:
        if line.strip() == 'Else':
            skip = False
        else:
            continue
            
    if line.strip() == 'On Error Resume Next' and lines[i+1].strip() == 'wsMain.Shapes("shpDemoLockG3G4").Delete' and lines[i+2].strip() == 'wsMain.Shapes("shpDemoMsgText").Delete':
        skip = True
        new_else_ui = """        On Error Resume Next
        wsMain.Shapes("shpDemoLockG3G4").Delete
        wsMain.Shapes("shpDemoMsgText").Delete
        wsMain.Range("E19:K20").ClearContents
        On Error GoTo 0\n"""
        new_lines.append(new_else_ui)
        continue
        
    if skip and line.strip() == 'On Error GoTo 0':
        skip = False
        continue
        
    if not skip:
        new_lines.append(line)

content = "".join(new_lines)
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_174"', 'Attribute VB_Name = "Goren_Claude_V2_175"')
content = content.replace('VERSION: V2.174', 'VERSION: V2.175')
content = content.replace('APP_VERSION As String = "2.174"', 'APP_VERSION As String = "2.175"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.175 correctly!")
