import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.110.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# I will add a simple InputBox for password in ToggleHiddenSheets
new_toggle = """Public Sub ToggleHiddenSheets()
    Dim pwd As String
    pwd = InputBoxU(ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & ":", ChrW(1488) & ChrW(1489) & ChrW(1496) & ChrW(1495) & ChrW(1492))
    If pwd <> "Z961814r" Then
        MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
"""
content = content.replace("Public Sub ToggleHiddenSheets()\n    Dim ws As Worksheet\n", new_toggle)

# Update version string just to be safe, V2.111
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_110"', 'Attribute VB_Name = "Goren_Claude_V2_111"')
content = content.replace('VERSION: V2.110', 'VERSION: V2.111')
content = content.replace('APP_VERSION As String = "2.110"', 'APP_VERSION As String = "2.111"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.111.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.111 created.")
