import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A1_RunDataValidation(" in line or "Public Sub RunDataValidation" in line:
        in_func = True
    if in_func:
        if i - lines.index("Public Sub A1_RunDataValidation()\n") > 200: break
        if "End Sub" in line:
            for j in range(max(0, i-20), i+2):
                print(f"[{j+1}] {lines[j].strip()}")
            break

