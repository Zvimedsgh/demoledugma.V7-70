import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_create = False
for i, line in enumerate(lines):
    if "Public Sub CreateInstructionSheets" in line:
        in_create = True
    if in_create and "H_SET_FIELDMAP" in line:
        for j in range(i, i+50):
            print(f"[{j+1}] {lines[j].strip()}")
        break

