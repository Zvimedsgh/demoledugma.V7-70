import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SendForReview()" in line:
        in_func = True
    elif in_func and "End Sub" in line:
        break
    elif in_func and "askSent = MsgBoxU" in line:
        for j in range(max(0, i-40), i):
            print(f"[{j+1}] {lines[j].strip()}")
        break

