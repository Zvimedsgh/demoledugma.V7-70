import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(1490, 1535):
    if "isDemoMode" in lines[i] or "yearVal" in lines[i] or "refYear" in lines[i]:
        print(f"[{i+1}] {lines[i].strip()}")

