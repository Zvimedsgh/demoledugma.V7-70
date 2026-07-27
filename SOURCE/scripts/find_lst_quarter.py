import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043_20260702_1545.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "lst_quarter" in line or "lst_month" in line:
        print(f"[{i}] {lines[j].strip() if 'j' in locals() else lines[i].strip()}")
