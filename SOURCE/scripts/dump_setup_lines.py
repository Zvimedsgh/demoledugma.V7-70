import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.078.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_setup = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub A00_SetupMainSheet_V2078()" in line_strip:
        in_setup = True
        print(f"[{i+1}] {line.strip()}")
    elif in_setup and "End Sub" in line_strip:
        print(f"[{i+1}] {line.strip()}")
        break
    elif in_setup and line_strip.startswith("1"):
        print(f"[{i+1}] {line.strip()}")

