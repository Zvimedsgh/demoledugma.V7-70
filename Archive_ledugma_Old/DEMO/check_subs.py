import re

with open(r'C:\ledugma\DEMO\modDemoReports_V9.40.bas', 'r', encoding='windows-1255', errors='ignore') as f:
    lines = f.readlines()

in_sub = False
sub_name = ""
sub_start_line = 0

for i, line in enumerate(lines):
    line_s = line.strip()
    if line_s.startswith("Public Sub ") or line_s.startswith("Private Sub ") or line_s.startswith("Sub ") or line_s.startswith("Public Function ") or line_s.startswith("Private Function ") or line_s.startswith("Function "):
        if in_sub:
            print(f"Error: Found new Sub/Function at line {i+1} ({line_s}) but previous Sub '{sub_name}' from line {sub_start_line} was not closed!")
        in_sub = True
        sub_name = line_s
        sub_start_line = i + 1
    elif line_s == "End Sub" or line_s == "End Function":
        if not in_sub:
            print(f"Error: Found End at line {i+1} but not inside a Sub!")
        in_sub = False

if in_sub:
    print(f"Error: File ended but Sub '{sub_name}' from line {sub_start_line} was never closed!")
