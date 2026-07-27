import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.172.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

start_idx = -1
for i, line in enumerate(lines):
    if "Public Sub FixUIButtons()" in line:
        start_idx = i
    if "End Sub" in line and start_idx != -1 and i > start_idx + 10:
        print(f"End Sub for FixUIButtons found at line {i+1}")
        break

