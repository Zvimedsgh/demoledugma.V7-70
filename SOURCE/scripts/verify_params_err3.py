import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3810, 3865):
    print(f"[{i}] {lines[i].strip()}")
