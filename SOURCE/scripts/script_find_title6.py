import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

lines = content.splitlines()
in_setup = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_setup = True
    if in_setup and "End Sub" in line:
        break
    if in_setup and "Delete" in line:
        print(f"Line {i+1}: {line}")
