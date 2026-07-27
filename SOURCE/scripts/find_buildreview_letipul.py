import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "If SheetExists(REVIEW_SHEET_NAME()) Then" in line:
        for j in range(i, i+10):
            print(f"[{j}] {lines[j].strip()}")
