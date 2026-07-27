import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "yearVal =" in line and "2025" in line:
        print(f"[{i}] {line.strip()}")
    if "refYear =" in line and "2024" in line:
        print(f"[{i}] {line.strip()}")
