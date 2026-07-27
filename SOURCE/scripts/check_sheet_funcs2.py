import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.062.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    line_strip = line.strip()
    if line_strip.startswith("Private Function H_") or line_strip.startswith("Private Function SHEET_"):
        for j in range(i, i+3):
            print(f"[{j+1}] {lines[j].strip()}")

