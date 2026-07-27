import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print(f"Total lines: {len(content.splitlines())}")
