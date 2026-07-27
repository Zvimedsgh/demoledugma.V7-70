import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.104.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Fallback to default if empty" in line:
        for j in range(i, min(len(lines), i+5)):
            print(f"[{j+1}] {lines[j].strip()}")
        break

