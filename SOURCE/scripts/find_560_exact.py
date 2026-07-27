import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    # Match any line that starts with "560" after trimming, followed by space
    line_trimmed = line.strip()
    if line_trimmed.startswith("560 ") or line_trimmed == "560":
        print(f"Line {i+1}: {line_trimmed}")
