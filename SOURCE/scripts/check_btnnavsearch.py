import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.135.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet" in line:
        in_func = True
    if in_func:
        if "btnNavSearch" in line:
            for j in range(max(0, i-5), min(len(lines), i+15)):
                print(f"[{j+1}] {lines[j].strip()}")
            break

