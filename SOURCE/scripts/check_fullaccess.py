import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "If accessLevel = fullAccess Or UCase$(accessLevel) = \"FULLACCESS\" Then" in line:
        in_func = True
        for j in range(max(0, i-2), min(len(lines), i+15)):
            print(f"[{j+1}] {lines[j].strip()}")
        break

