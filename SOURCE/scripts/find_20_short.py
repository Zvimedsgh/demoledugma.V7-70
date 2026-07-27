import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "MIN_PREMIUM" in line or "20" in line:
        if "20" in line and len(line) < 100:
            print(f"[{i}] {line.strip()}")
