import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038_20260702_1440.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "GetActiveAgencyName" in line or "Title" in line or "כותרת" in line or "C1" in line:
        if i > 3700 and i < 3800:
            print(f"[{i}] {lines[i].strip()}")
