import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.031.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub NavSettings_FieldMap" in line:
        print(f"[{i}] {line.strip()}")
        for j in range(i, i+5):
            print(f"  [{j}] {lines[j].strip()}")
