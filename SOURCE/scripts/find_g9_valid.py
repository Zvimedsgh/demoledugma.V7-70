import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043_20260702_1545.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Validation.Add" in line and ("G9" in line or "G10" in line):
        for j in range(i-2, i+5):
            print(f"[{j}] {lines[j].strip()}")
        break
