import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation()" in line:
        in_func = True
    if in_func:
        if "1050" in line or "1060" in line or "1070" in line or "1080" in line or "ERR_HANDLER:" in line:
            for j in range(max(0, i-5), i+2):
                print(f"[{j+1}] {lines[j].strip()}")

