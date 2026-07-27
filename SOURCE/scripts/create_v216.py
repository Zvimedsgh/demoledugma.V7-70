import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.215.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.216.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_hide = """    If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then hideIt = False
    End If"""

new_hide = """    Dim wsName1 As String, wsName2 As String
    wsName1 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    wsName2 = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)
    
    If ws.Name = wsName1 Or ws.Name = wsName2 Or ws.Name = "Sheet2" Then
        If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then
            hideIt = False
            If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
            On Error Resume Next
            ws.Move After:=ThisWorkbook.Worksheets(ctrlName)
            On Error GoTo 0
        End If
    End If"""

# Python string matching might fail due to comments, let's use a robust way.
import re
pattern = re.compile(r'If ws\.Name = ChrW\(1492\).*?Then.*?If UCase\$\(ThisWorkbook.*?hideIt = False\s*End If', re.DOTALL)
content = pattern.sub(new_hide, content, count=1)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_215"', 'Attribute VB_Name = "Goren_Claude_V2_216"')
content = content.replace('VERSION: V2.215', 'VERSION: V2.216')
content = content.replace('APP_VERSION As String = "2.215"', 'APP_VERSION As String = "2.216"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.216")
