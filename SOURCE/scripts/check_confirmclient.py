import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.148.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ConfirmClientSelection(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 100: break
        print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line:
            break

