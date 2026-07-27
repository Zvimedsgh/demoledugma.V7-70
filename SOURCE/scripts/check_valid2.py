import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.114.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Validation.Add" in line or "Validation.Modify" in line or "Formula1:=" in line:
        print(f"[{i+1}] {line.strip()}")

