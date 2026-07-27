import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.044.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "listName = \"=lst_half_year\"" in line or "ThisWorkbook.Names.Add \"lst_temp_filter\"" in line:
        print(f"[{i}] {lines[i].strip()}")
