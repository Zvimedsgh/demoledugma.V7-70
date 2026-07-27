import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.102.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_func = True
        print(f"[{i+1}] {line.strip()}")
    elif in_func and "End Sub" in line:
        print(f"[{i+1}] {line.strip()}")
        break
    elif in_func:
        if "Range" in line and "Select" in line or "GoTo" in line or "Goto" in line or "AgencyName" in line or "ERR_HANDLER" in line:
            print(f"[{i+1}] {line.strip()}")

