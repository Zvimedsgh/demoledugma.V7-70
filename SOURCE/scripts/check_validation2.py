import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "rngClientName" in line and "Validation.Add" in line:
        for j in range(max(0, i-5), min(len(lines), i+2)):
            print(f"[{j+1}] {lines[j].strip()}")

