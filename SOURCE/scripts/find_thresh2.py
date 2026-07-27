import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "PREMIUM_THRESHOLD" in line or "20" in line:
        if "PREMIUM_THRESHOLD" in line and "Value" in line:
            print(f"[{i}] {line.strip()}")
