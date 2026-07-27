import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "FORCE_DEMO_MODE" in line or "wsMain.Range(\"A1:Z1\").UnMerge" in line or "wsMain.Range(\"G1:P1\").Merge" in line:
        print(f"[{i}] {lines[i].strip()}")
