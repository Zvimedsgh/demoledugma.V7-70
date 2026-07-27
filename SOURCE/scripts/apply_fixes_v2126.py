import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.125.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    # 1. Add explanation in SearchClientName
    if "' Activate the search sheet and put cursor in search cell" in line:
        new_lines.append("""' Add explanation text
wsSearch.Cells(5, 2).Value = ChrW(1492) & ChrW(1505) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & ":"
wsSearch.Cells(6, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1512) & ChrW(1511) & " 77"
wsSearch.Cells(7, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"
wsSearch.Cells(8, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"
wsSearch.Cells(9, 2).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " ???: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1501) & " 3 " & ChrW(1514) & ChrW(1493) & ChrW(1493) & ChrW(1497) & ChrW(1501)
wsSearch.Range("B5:B9").Font.Size = 11
wsSearch.Range("B5:B9").Font.Color = RGB(100, 100, 100)
wsSearch.Cells(5, 2).Font.Bold = True

""")
        new_lines.append(line)
        continue
    
    # 2. Add Demo Message in ApplyDemoLockOnOpen
    if 'shpLock.OnAction = "DemoModeRestricted"' in line:
        new_lines.append(line)
        new_lines.append("""
On Error Resume Next
wsMain.Shapes("shpDemoMsgText").Delete
On Error GoTo 0
Dim shpDemoMsg As Shape
Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, wsMain.Range("D3").Left, wsMain.Range("D3").Top, 350, 40)
shpDemoMsg.Name = "shpDemoMsgText"
shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 16
shpDemoMsg.TextFrame2.TextRange.Font.Bold = -1
shpDemoMsg.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 255)
shpDemoMsg.Line.Visible = 0
shpDemoMsg.Fill.Visible = 0
""")
        continue
        
    # 3. Add delete in ApplyDemoLockOnOpen (Else branch)
    if 'wsMain.Shapes("shpDemoLockG3G4").Delete' in line and "ApplyDemoLockOnOpen" not in "".join(new_lines[-100:]):
        # wait, we must make sure it's the right one.
        pass

    new_lines.append(line)

content = "".join(new_lines)

# Manual target delete for the Else branch in ApplyDemoLockOnOpen
target_lock_else = """wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
On Error Resume Next
wsMain.Shapes("shpDemoLockG3G4").Delete
On Error GoTo 0"""
new_lock_else = """wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
On Error Resume Next
wsMain.Shapes("shpDemoLockG3G4").Delete
wsMain.Shapes("shpDemoMsgText").Delete
On Error GoTo 0"""

if target_lock_else in content:
    content = content.replace(target_lock_else, new_lock_else)
    print("Replaced lock else")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_125"', 'Attribute VB_Name = "Goren_Claude_V2_126"')
content = content.replace('VERSION: V2.125', 'VERSION: V2.126')
content = content.replace('APP_VERSION As String = "2.125"', 'APP_VERSION As String = "2.126"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.126.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.126 created.")
