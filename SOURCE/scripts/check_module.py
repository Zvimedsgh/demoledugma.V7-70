import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

vars = set()
for i, line in enumerate(lines):
    line_strip = line.strip()
    if line_strip.startswith("Private Sub") or line_strip.startswith("Public Sub") or line_strip.startswith("Function"):
        break
    if line_strip.startswith("Private ") or line_strip.startswith("Public ") or line_strip.startswith("Dim "):
        if "Const " in line_strip:
            var_name = line_strip.split("Const ")[1].split()[0]
        else:
            dim_str = line_strip.replace("Private ", "").replace("Public ", "").replace("Dim ", "").replace("WithEvents ", "")
            var_name = dim_str.split()[0].replace("()", "").split(",")[0]
            
        if var_name.lower() in vars:
            print(f"DUPLICATE MODULE LEVEL: {var_name}")
        vars.add(var_name.lower())

