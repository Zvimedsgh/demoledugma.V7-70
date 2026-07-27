import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "200000" in line:
        print(f"[{i+1}] {line.strip()}")
    if "SetupMainSheet" in line:
        for j in range(i, i+10):
            print(f"[{j+1}] {lines[j].strip()}")

