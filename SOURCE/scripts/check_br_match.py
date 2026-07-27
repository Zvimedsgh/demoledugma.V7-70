import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.088.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "matchBr = Application.Match" in line:
        print(f"[{i+1}] {line.strip()}")

