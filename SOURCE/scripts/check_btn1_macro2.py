import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "OnAction" in line and "btn1" in line.lower() or "btnStep1" in line:
        print(f"[{i+1}] {line.strip()}")

