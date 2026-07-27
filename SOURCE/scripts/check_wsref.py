import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 1000: break
        if "Set wsRef =" in line or "refYear" in line:
            for j in range(max(0, i-2), min(len(lines), i+3)):
                print(f"[{j+1}] {lines[j].strip()}")
            print("-" * 40)

