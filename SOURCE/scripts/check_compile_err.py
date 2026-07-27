import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Sub InitRawColumns" in line:
        print(f"InitRawColumns starts at {i+1}")
    if "Private Const BASE_COL_ID" in line:
        print(f"BASE_COL_ID starts at {i+1}")
        break

