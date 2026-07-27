import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

count = 0
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet" in line:
        print(f"Found at {i}")
        count += 1
print(f"Total: {count}")
