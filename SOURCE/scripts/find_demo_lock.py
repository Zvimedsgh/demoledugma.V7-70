import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.039_20260702_1445.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub SaveReportsToFolder" in line or "Public Sub ViewReportsFolder" in line or "Public Sub NewClients" in line:
        for j in range(i, i+10):
            print(f"[{j}] {lines[j].strip()}")
