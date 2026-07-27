import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def check_sub(sub_name):
    in_sub = False
    dims = []
    for i, line in enumerate(lines):
        if f"Public Sub {sub_name}" in line:
            in_sub = True
        if in_sub and "End Sub" in line:
            break
        if in_sub and "Dim " in line:
            parts = line.strip().split("Dim ")[1].split(",")
            for p in parts:
                var_name = p.strip().split(" ")[0]
                dims.append((i, var_name))

    seen = {}
    for i, var_name in dims:
        if var_name in seen:
            print(f"DUPLICATE FOUND in {sub_name}: {var_name} at line {i} (previously at {seen[var_name]})")
        else:
            seen[var_name] = i

check_sub("UpdatePeriodDropdown")
check_sub("UpdateFilterValueDropdown")
