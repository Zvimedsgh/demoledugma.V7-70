import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.085.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub UpdateClientList()" in line_strip:
        in_func = True
        print(f"[{i+1}] {line.strip()}")
    elif in_func and "End Sub" in line_strip:
        print(f"[{i+1}] {line.strip()}")
        break
    elif in_func:
        print(f"[{i+1}] {line.strip()}")

