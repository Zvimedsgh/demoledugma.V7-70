import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.165.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
end_idx = -1

for i, line in enumerate(lines):
    if 'wsMain.Shapes("shpDemoMsgText").Delete' in line and start_idx == -1:
        # found the first one in the IF block
        start_idx = i
    if '1750 Else' in line and start_idx != -1:
        end_idx = i
        break

if start_idx != -1 and end_idx != -1:
    new_lines = lines[:start_idx]
    
    new_code = """    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo ERR_HANDLER
    With wsMain.Range("E19:K20")
        .Merge
        .Value = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ": 054-6677396"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 255)
    End With
"""
    new_lines.append(new_code)
    new_lines.extend(lines[end_idx:])
    
    # Update version
    for i in range(len(new_lines)):
        new_lines[i] = new_lines[i].replace('Attribute VB_Name = "Goren_Claude_V2_165"', 'Attribute VB_Name = "Goren_Claude_V2_166"')
        new_lines[i] = new_lines[i].replace('VERSION: V2.165', 'VERSION: V2.166')
        new_lines[i] = new_lines[i].replace('APP_VERSION As String = "2.165"', 'APP_VERSION As String = "2.166"')
        
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.166.bas', 'w', encoding='utf-8') as f:
        f.writelines(new_lines)
    print("Created V2.166 properly replacing the shape code.")
else:
    print("Could not find the block to replace!")
