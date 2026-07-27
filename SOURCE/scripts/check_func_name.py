import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.150.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(1518, -1, -1):
    if "Sub " in lines[i] or "Function " in lines[i]:
        print(f"[{i+1}] {lines[i].strip()}")
        break

