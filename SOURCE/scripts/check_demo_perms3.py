import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.205.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub CheckUserPermissions" in line:
        for j in range(i+60, i+120):
            try:
                print(f"[{j+1}] {lines[j].strip()}")
            except:
                pass
        break
