import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1507.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        for j in range(i+50, i+100):
            print(f"[{j}] {lines[j].strip()}")
        break
