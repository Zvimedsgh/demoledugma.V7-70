import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "rngClientName" in line:
        for j in range(i-2, i+3):
            print(f"[{j+1}] {lines[j].strip()}")

