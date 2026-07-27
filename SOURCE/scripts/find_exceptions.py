import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "H_EXCEPTIONS" in line or "1495" in line:  # 1495 is Chet
        if "Function H_EXCEPTIONS" in line:
            print(f"[{i}] {line.strip()}")
