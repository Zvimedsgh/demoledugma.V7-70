import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.146.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = """    Else
        ' Showing sheets - require password
        Dim pwd As String
        pwd = InputBox(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))
        If pwd <> "Z961814r" Then
            MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbCritical
            Exit Sub
        End If
        ShowHiddenSheets
    End If"""

new = """    Else
        ' Showing sheets - temporarily disabled password for debugging
        ShowHiddenSheets
    End If"""

if target in content:
    content = content.replace(target, new)
    print("Fixed ToggleHiddenSheets password.")
else:
    print("Could not find ToggleHiddenSheets target.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_146"', 'Attribute VB_Name = "Goren_Claude_V2_147"')
content = content.replace('VERSION: V2.146', 'VERSION: V2.147')
content = content.replace('APP_VERSION As String = "2.146"', 'APP_VERSION As String = "2.147"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.147.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.147 created.")
