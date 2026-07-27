import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041_20260702_1520.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Exit Sub" in line and "CheckUserPermissions" in lines[i+2]:
        for j in range(i-5, i+5):
            print(f"[{j}] {lines[j].strip()}")
        break
