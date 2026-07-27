import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.069.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub A00_SetupMainSheet()" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        break
    elif in_func and ("H2" in line_strip or "F2" in line_strip or "G2" in line_strip or "I2" in line_strip):
        print(f"[{i+1}] {line_strip}")

