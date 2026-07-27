import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
vars = set()
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub BuildPresentation()" in line_strip:
        in_func = True
        vars = set()
    elif in_func and "End Sub" in line_strip:
        break
    elif in_func and line_strip.startswith("Dim "):
        dim_str = line_strip[4:]
        parts = dim_str.split(",")
        for part in parts:
            part = part.strip()
            var_name = part.split()[0].replace("()", "")
            if var_name.lower() in vars:
                print(f"DUPLICATE FOUND IN BuildPresentation: {var_name}")
            vars.add(var_name.lower())

