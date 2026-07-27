import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "btnHafad" in line or "btnKulam" in line or "חפ""ד" in line or "כולם" in line:
        print(f"[{i+1}] {line.strip()}")

