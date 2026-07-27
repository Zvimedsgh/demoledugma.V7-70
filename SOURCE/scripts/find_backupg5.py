import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "backupG5" in line:
        for j in range(i-20, i+20):
            if j >= 0 and j < len(lines):
                print(f"[{j}] {lines[j].strip()}")
        break
