import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.075.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_setup = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub A00_SetupMainSheet()" in line_strip:
        in_setup = True
    elif in_setup and "End Sub" in line_strip:
        break
    elif in_setup and "1731 On Error GoTo ERR_HANDLER" in line_strip:
        for j in range(i, i+15):
            print(f"[{j+1}] {lines[j].strip()}")
        break

