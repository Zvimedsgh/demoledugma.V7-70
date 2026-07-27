import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.043_20260702_1545.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3530, 3545):
    print(f"[{i}] {lines[i].strip()}")
