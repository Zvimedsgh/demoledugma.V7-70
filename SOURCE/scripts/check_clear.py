import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "wsSrc.Cells.Clear" in line or "wsSrc.Range(" in line and "Clear" in line:
        print(f"[{i}] {lines[i].strip()}")
    if "wsRef.Cells.Clear" in line or "wsRef.Range(" in line and "Clear" in line:
        print(f"[{i}] {lines[i].strip()}")

