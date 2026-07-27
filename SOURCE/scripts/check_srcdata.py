import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.097.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "srcData =" in line:
        for j in range(i-5, min(len(lines), i+5)):
            print(f"[{j+1}] {lines[j].strip()}")
        break

