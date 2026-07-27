import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_setup = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_setup = True
    if in_setup and "End Sub" in line:
        in_setup = False
    
    if in_setup and "H_SET_PARAMS" in line:
        print("FOUND BLOCK AT", i)
        for j in range(i, i+15):
            print(f"[{j}] {lines[j].strip()}")
        break
