import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.105.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Line.Visible =" in line or "Line.ForeColor" in line or "Border" in line:
        if "RGB(255, 0, 0)" in line or "RGB(200, 0, 0)" in line or "msoTrue" in line:
            print(f"[{i+1}] {line.strip()}")

