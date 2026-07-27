import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Const BASE_COL_CUSTNAME As Long =" in line:
        print(f"[{i+1}] {line.strip()}")
    if "Private Const RAW_CUSTNAME As Long =" in line:
        print(f"[{i+1}] {line.strip()}")

