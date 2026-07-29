import re

filename = r"C:\ledugma\DEMO\modDemoReports_V9.34.bas"
with open(filename, "r", encoding="windows-1255", errors="ignore") as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "1502" in line and "1489" in line and "1510" in line: # מבצע
        print(f"Line {i+1}: {line.strip()}")
    elif "1506" in line and "1497" in line and "1489" in line: # עיבוד
        print(f"Line {i+1}: {line.strip()}")
    elif "1502" in line and "1510" in line and "1490" in line: # מצגת
        print(f"Line {i+1}: {line.strip()}")
