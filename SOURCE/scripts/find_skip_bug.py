import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "If isDemoMode Then" in line and "BuildReview" in "".join(lines[max(0, i-50):i]):
        for j in range(i, i+20):
            print(f"[{j}] {lines[j].strip()}")
        break
