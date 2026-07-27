import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "AddTextbox" in line and "shpDemoMsg" in line:
        print(f"[{i+1}] {lines[i].strip()}")

