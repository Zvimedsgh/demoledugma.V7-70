import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.152.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = """    If wb Is Nothing Then
        Application.DisplayAlerts = True
        On Error Resume Next
        Workbooks.Open fPath, ReadOnly:=True
        On Error GoTo 0
        Application.DisplayAlerts = False
    End If"""

new_code = """    If wb Is Nothing Then
        Application.DisplayAlerts = False ' Suppress KuTools or Add-In errors
        On Error Resume Next
        Workbooks.Open fPath, ReadOnly:=True
        On Error GoTo 0
        Application.DisplayAlerts = True ' Restore
    End If"""

if target in content:
    content = content.replace(target, new_code)
    print("Fixed SafeOpenWorkbook DisplayAlerts.")
else:
    print("Could not find target string.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_152"', 'Attribute VB_Name = "Goren_Claude_V2_153"')
content = content.replace('VERSION: V2.152', 'VERSION: V2.153')
content = content.replace('APP_VERSION As String = "2.152"', 'APP_VERSION As String = "2.153"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.153 created.")
