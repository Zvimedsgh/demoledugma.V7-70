import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_search = """For r = 2 To lastRow
cName = Trim$(CStr(ws.Cells(r, 6).Value2))
If cName <> "" Then
If InStr(1, cName, searchText, vbTextCompare) > 0 Then
If Not dict.Exists(cName) Then dict.Add cName, 1
End If
End If
Next r"""

new_search = """Dim bWildcard As Boolean
bWildcard = (InStr(1, searchText, "*") > 0 Or InStr(1, searchText, "?") > 0)

For r = 2 To lastRow
cName = Trim$(CStr(ws.Cells(r, 6).Value2))
If cName <> "" Then
    Dim bMatch As Boolean
    bMatch = False
    If bWildcard Then
        If UCase(cName) Like UCase(searchText) Then bMatch = True
    Else
        If UCase(cName) = UCase(searchText) Then bMatch = True
    End If
    
    If bMatch Then
        If Not dict.Exists(cName) Then dict.Add cName, 1
    End If
End If
Next r"""

if target_search in content:
    content = content.replace(target_search, new_search)
    print("Replaced target_search")
else:
    print("Could not find target_search")
    idx = content.find("For r = 2 To lastRow")
    if idx != -1:
        print("Found For r = 2 To lastRow, maybe whitespace issue.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_124"', 'Attribute VB_Name = "Goren_Claude_V2_125"')
content = content.replace('VERSION: V2.124', 'VERSION: V2.125')
content = content.replace('APP_VERSION As String = "2.124"', 'APP_VERSION As String = "2.125"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.125.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.125 created.")
