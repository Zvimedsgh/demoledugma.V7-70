import os
import re

file_path = r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()
for i, line in enumerate(lines):
    if re.search(r'LEVAV|לבב', line, re.IGNORECASE):
        print(f"{i+1}: {line.strip()}")
