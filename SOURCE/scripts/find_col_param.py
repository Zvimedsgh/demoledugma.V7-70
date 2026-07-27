import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "COL_PARAM_NAME" in line:
        print(f"[{i}] {lines[i].strip()}")
        break
