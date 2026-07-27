import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "H_BASE" in line and "ws." in line or "Name" in line and "H_BASE" in line or "Add" in line and "H_BASE" in line or "Copy" in line and "H_BASE" in line:
        for j in range(max(0, i-2), min(len(lines), i+3)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 40)

