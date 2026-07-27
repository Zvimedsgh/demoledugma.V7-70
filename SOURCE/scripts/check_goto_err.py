import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.071.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "GoTo ERR_HANDLER" in line and "On Error GoTo" not in line:
        print(f"[{i+1}] {line.strip()}")

