import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.174.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub CheckUserPermissions()" in line:
        print(f"CheckUserPermissions starts at {i+1} in V2.174")

