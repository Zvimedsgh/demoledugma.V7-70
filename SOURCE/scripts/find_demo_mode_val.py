import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1507.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "DEMO_MODE" in line and "Value =" in line:
        for j in range(i-2, i+5):
            print(f"[{j}] {lines[j].strip()}")
        break
