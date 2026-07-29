import re

filename = r"C:\ledugma\DEMO\modDemoReports_V9.34.bas"
with open(filename, "r", encoding="windows-1255", errors="ignore") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    # if re.search(r'Cells\(\s*15', line) or re.search(r'Cells\(\s*17', line):
    if "Cells(15" in line or "Cells(17" in line or "Cells(18" in line or "Cells(14" in line:
        print(f"Line {i+1}: {line.strip()}")
