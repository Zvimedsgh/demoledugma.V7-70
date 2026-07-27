import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.131.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Exit" in line and "btn" in line:
        print(f"[{i+1}] {lines[i].strip()}")

