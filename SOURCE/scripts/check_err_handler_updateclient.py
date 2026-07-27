import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub UpdateClientList(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 150: break
        if "ERR_HANDLER:" in line:
            print(f"Found ERR_HANDLER at line {i+1}")
        if "End Sub" in line:
            print(f"End of sub at line {i+1}")
            break

