import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

current_func = ""
vars_in_func = set()

for i, line in enumerate(lines):
    line_strip = line.strip()
    
    # Check for function/sub start
    m = re.match(r'^(?:Private |Public |)(?:Sub |Function |Property Get |Property Let )([A-Za-z0-9_]+)', line_strip)
    if m:
        current_func = m.group(1)
        vars_in_func = set()
        continue
        
    if "End Sub" in line_strip or "End Function" in line_strip or "End Property" in line_strip:
        current_func = ""
        vars_in_func = set()
        continue
        
    if current_func and line_strip.startswith("Dim "):
        # e.g. Dim myVar As String, myVar2 As Long
        # Need to parse comma separated
        # Strip "Dim "
        dim_str = line_strip[4:]
        # Split by comma (ignoring commas in strings or arrays, but usually variables are simple)
        parts = dim_str.split(",")
        for part in parts:
            part = part.strip()
            # The variable name is the first word before " As " or "()"
            var_name = part.split()[0].replace("()", "")
            if var_name.lower() in vars_in_func:
                print(f"DUPLICATE in {current_func}: {var_name} at line {i+1}")
            vars_in_func.add(var_name.lower())

