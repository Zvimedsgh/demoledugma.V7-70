import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_pres = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation()" in line:
        in_pres = True
    if in_pres and ("WindowState" in line or "Show" in line or "View" in line or "FullScreen" in line):
        print(f"[{i+1}] {line.strip()}")
    if in_pres and "End Sub" in line:
        break

