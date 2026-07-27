import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ResetHomeDefaults()" in line or "Private Sub ResetHomeDefaults()" in line:
        in_func = True
    if in_func and "End Sub" in line:
        for j in range(max(0, i-5), min(len(lines), i+5)):
            print(f"[{j}] {lines[j].strip()}")
        break
    if in_func:
        if "On Error" in line or "Err.Clear" in line or "ERR_HANDLER:" in line:
            print(f"[{i}] {lines[i].strip()}")
