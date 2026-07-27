import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "ExecuteExcel4Macro" in line or "CommandBars" in line or "DisplayFormulaBar" in line or "DisplayHeadings" in line or "DisplayWorkbookTabs" in line:
        print(f"[{i+1}] {line.strip()}")

