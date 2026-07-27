import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "BuildSummarySheet" in line and "Sub" not in line:
        print(f"[{i+1}] {line.strip()}")

