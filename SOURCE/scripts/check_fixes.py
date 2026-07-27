import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def print_func(func_name):
    in_func = False
    for i, line in enumerate(lines):
        if func_name in line and ("Sub " in line or "Function " in line):
            in_func = True
            print(f"[{i+1}] {line.strip()}")
        elif in_func and "End Sub" in line or in_func and "End Function" in line:
            print(f"[{i+1}] {line.strip()}")
            break
        elif in_func:
            print(f"[{i+1}] {line.strip()}")
    print("-" * 20)

print_func("UpdateFilterValueDropdown")
print_func("BuildSummarySheet")

def get_func(lines, idx):
    for i in range(idx, -1, -1):
        m = re.search(r'(Sub|Function)\s+([A-Za-z0-9_]+)', lines[i])
        if m: return m.group(2)
    return ""

for i, line in enumerate(lines):
    if ".Delete" in line and "ws" in line and "Presentation" in get_func(lines, i):
        for j in range(max(0, i-2), min(len(lines), i+3)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

