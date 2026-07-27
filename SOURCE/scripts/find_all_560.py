import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if line.strip().startswith("560 "):
        for j in range(i-5, i+5):
            print(f"[{j}] {lines[j].strip()}")
        print("-" * 20)
