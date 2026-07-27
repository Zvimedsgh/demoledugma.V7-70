import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.140.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "OnKey" in line or "Worksheet_Change" in line:
        print(f"[{i+1}] {line.strip()}")

