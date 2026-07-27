import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "CreatePowerPoint" in line or "Sub CreatePPT" in line or "Miktze" in line.lower() or "powerpoint" in line.lower():
        print(f"[{i+1}] {line.strip()}")

