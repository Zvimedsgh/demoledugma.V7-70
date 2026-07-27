import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.060.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "\\u05d0\\u05e8\\u05d2\\u05d5\\u05df \\u05de\\u05d7\\u05d3\\u05e9" in line.encode('unicode_escape').decode() or "ארגון" in line:
        print(f"[{i+1}] {line.strip()}")

