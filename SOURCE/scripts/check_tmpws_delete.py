import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "tmpWs.Delete" in line or "tmpWs" in line and "Delete" in line:
        for j in range(i-2, i+3):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

