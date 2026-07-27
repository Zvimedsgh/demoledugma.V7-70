import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 1000: break
        if "lst_clients" in line or "Validation" in line:
            print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line and not "Setup" in line:
            break

