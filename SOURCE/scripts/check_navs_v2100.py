import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.100.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for func in ["Public Sub NavToIndex()", "Public Sub NavSettings_Home()", "Public Sub NavSettings_Menu()"]:
    in_func = False
    for i, line in enumerate(lines):
        if func in line:
            in_func = True
            print(f"[{i+1}] {line.strip()}")
        elif in_func and "End Sub" in line:
            print(f"[{i+1}] {line.strip()}")
            print("-" * 20)
            break
        elif in_func:
            print(f"[{i+1}] {line.strip()}")

