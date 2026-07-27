import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "RAW_" in line and "Private Const" not in line and "InitRawColumns" not in line:
        print(f"[{i+1}] {line.strip()}")

