import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.062.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    line_strip = line.strip()
    if line_strip.startswith("Private Function ") and "SHEET_NAME" in line_strip:
        for j in range(i, i+5):
            print(f"[{j+1}] {lines[j].strip()}")

