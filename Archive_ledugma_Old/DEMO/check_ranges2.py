import re

with open(r'C:\ledugma\DEMO\modDemoReports_V9.36.bas', encoding='windows-1255', errors='ignore') as f:
    lines = f.readlines()

current_macro = ""
for line in lines:
    if line.startswith("Public Sub") or line.startswith("Private Sub") or line.startswith("Function"):
        current_macro = line.strip()
    
    m = re.findall(r'(?:wsMain|ThisWorkbook\.Worksheets\(CONTROL_SHEET_NAME\(\)\))\.Range\("([A-Z][0-9]+(?::[A-Z][0-9]+)?)"\)', line)
    if m:
        if "SetupMainSheet" not in current_macro:
            print(f"{current_macro} -> {m}")
