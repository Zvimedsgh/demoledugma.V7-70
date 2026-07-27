import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.061.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "isDemoMode" in line_strip and ("Build" in line_strip or "Export" in line_strip or "Review" in line_strip):
        print(f"[{i+1}] {line_strip}")
    elif "isDemoMode" in line_strip:
        pass # print(f"[{i+1}] {line_strip}")

