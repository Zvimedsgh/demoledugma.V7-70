import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "COL_AGENT" in line or "COL_TELLER" in line or "COL_COMPANY" in line or "COL_BRANCH" in line:
        print(f"[{i+1}] {line.strip()}")

