import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet" in line:
        in_func = True
    if in_func:
        if "shpDemoMsgText" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            break

