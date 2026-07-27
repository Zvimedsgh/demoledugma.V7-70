import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.097.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "720 " in line:
        for j in range(i-20, i+20):
            print(f"[{j+1}] {lines[j].strip()}")
        break

