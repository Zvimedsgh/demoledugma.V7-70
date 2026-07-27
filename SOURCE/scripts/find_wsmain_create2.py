import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044_20260702_1555.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "If wsMain Is Nothing Then" in line:
        for j in range(i-2, i+8):
            print(f"[{j}] {lines[j].strip()}")
        break
