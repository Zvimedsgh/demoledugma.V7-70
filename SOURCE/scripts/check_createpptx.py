import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def check_func(lines, func_name):
    in_func = False
    vars = set()
    for i, line in enumerate(lines):
        line_strip = line.strip()
        if func_name in line_strip and ("Sub" in line_strip or "Function" in line_strip):
            in_func = True
        elif in_func and ("End Sub" in line_strip or "End Function" in line_strip):
            break
        elif in_func and "Dim " in line_strip:
            dim_str = line_strip.split("Dim ")[1]
            parts = dim_str.split(",")
            for part in parts:
                var_name = part.strip().split()[0].replace("()", "")
                if var_name.lower() in vars:
                    print(f"DUPLICATE in {func_name}: {var_name}")
                vars.add(var_name.lower())

check_func(lines, "CreatePPTX")

