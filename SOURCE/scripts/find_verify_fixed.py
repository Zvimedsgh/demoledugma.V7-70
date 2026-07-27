import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.049_20260702_1635.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "isDemoMode = FORCE_DEMO_MODE" in line:
        for j in range(i, i+5):
            print(f"[{j}] {lines[j].strip()}")
        break
