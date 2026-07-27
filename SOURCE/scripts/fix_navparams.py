import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_str = """Public Sub NavToParams()
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())"""

new_str = """Public Sub NavToParams()
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(H_SET_PARAMS())"""

if old_str in content:
    content = content.replace(old_str, new_str)
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Fixed NavToParams")
else:
    print("Could not find string")
