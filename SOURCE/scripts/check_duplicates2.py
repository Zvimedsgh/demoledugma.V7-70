import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

current_func = ""
vars_in_func = set()

for i, line in enumerate(lines):
    line_strip = line.strip()
    
    # Strip line numbers from the beginning if present
    m_ln = re.match(r'^\d+\s+(.*)', line_strip)
    if m_ln:
        line_strip = m_ln.group(1)
    
    # Check for function/sub start
    m = re.match(r'^(?:Private |Public |)(?:Sub |Function |Property Get |Property Let )([A-Za-z0-9_]+)\((.*?)\)', line_strip)
    if m:
        current_func = m.group(1)
        args_str = m.group(2)
        vars_in_func = set()
        
        # Parse args
        if args_str:
            args = args_str.split(",")
            for arg in args:
                arg = arg.strip()
                if arg.startswith("ByVal ") or arg.startswith("ByRef "):
                    arg = arg[6:].strip()
                if arg.startswith("Optional "):
                    arg = arg[9:].strip()
                if arg.startswith("ParamArray "):
                    arg = arg[11:].strip()
                var_name = arg.split()[0].replace("()", "")
                vars_in_func.add(var_name.lower())
        continue
        
    if "End Sub" in line_strip or "End Function" in line_strip or "End Property" in line_strip:
        current_func = ""
        vars_in_func = set()
        continue
        
    if current_func and line_strip.startswith("Dim "):
        dim_str = line_strip[4:]
        parts = dim_str.split(",")
        for part in parts:
            part = part.strip()
            var_name = part.split()[0].replace("()", "")
            if var_name.lower() in vars_in_func:
                print(f"DUPLICATE in {current_func}: {var_name} at line {i+1}")
            vars_in_func.add(var_name.lower())
            
    if current_func and line_strip.startswith("Const "):
        const_str = line_strip[6:]
        var_name = const_str.split()[0]
        if var_name.lower() in vars_in_func:
            print(f"DUPLICATE in {current_func}: {var_name} at line {i+1}")
        vars_in_func.add(var_name.lower())

