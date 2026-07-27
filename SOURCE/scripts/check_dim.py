import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub ExportCompCharts" in line:
        in_func = True
    elif "Private Sub ExportAgentCommissions" in line:
        in_func = True
    
    if in_func:
        if "End Sub" in line:
            in_func = False
            
        if "Dim " in line:
            print(f"[{i}] {line.strip()}")

