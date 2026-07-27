import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "APP_VERSION" in line:
        print(f"{lines[i].strip()}")

