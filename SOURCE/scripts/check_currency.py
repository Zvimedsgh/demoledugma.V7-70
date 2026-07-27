import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.085.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Function GetCurrencyRate" in line_strip:
        in_func = True
    elif in_func and "End Function" in line_strip:
        break
    elif in_func and "While " in line_strip or in_func and "For " in line_strip or in_func and "Do " in line_strip:
        print(f"[{i+1}] {line_strip}")

