import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "G10" in line or "H10" in line or "כולם" in line or "חפ" in line:
        if "Function" not in line and "MsgBoxU" not in line:
            print(f"[{i+1}] {line.strip()}")

