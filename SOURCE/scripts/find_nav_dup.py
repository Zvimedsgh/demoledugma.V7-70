import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

count = 0
for i, line in enumerate(lines):
    if "NavSettings_FieldMap" in line:
        print(f"[{i}] {line.strip()}")
        count += 1
print(f"Total occurrences: {count}")
