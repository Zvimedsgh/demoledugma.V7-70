import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(4810, 4890):
    if lines[i].strip().startswith("620"):
        print(f"[{i}] {lines[i].strip()}")
        break
