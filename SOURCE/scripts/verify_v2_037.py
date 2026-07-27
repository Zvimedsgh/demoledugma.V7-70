import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3500, 3540):
    if "isDemoMode" in lines[i] or "demoParam" in lines[i]:
        print(f"[{i}] {lines[i].strip()}")

for i in range(3810, 3880):
    print(f"[{i}] {lines[i].strip()}")
