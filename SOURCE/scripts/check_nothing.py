import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Set ppPres = Nothing" in line or "Set ppApp = Nothing" in line:
        print(f"[{i+1}] {line.strip()}")

