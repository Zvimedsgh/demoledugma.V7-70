import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.102.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Protect Password" in line:
        for j in range(i-2, i+3):
            print(f"[{j+1}] {lines[j].strip()}")

