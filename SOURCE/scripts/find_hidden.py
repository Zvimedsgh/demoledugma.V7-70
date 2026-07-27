import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsTemp.Visible" in line or "VeryHidden" in line:
        for j in range(i-2, i+2):
            print(f"[{j}] {lines[j].strip()}")
