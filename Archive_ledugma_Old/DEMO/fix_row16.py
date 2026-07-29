import os
import glob

filenames = glob.glob(r"C:\ledugma\DEMO\*.bas") + glob.glob(r"C:\ledugma\DEMO\*.cls")
found = False
for filename in filenames:
    try:
        with open(filename, "r", encoding="windows-1255", errors="ignore") as f:
            lines = f.readlines()
        
        for i, line in enumerate(lines):
            if "מבצע" in line or "מצגת" in line or "עיבוד" in line:
                print(f"File: {filename}, Line {i+1}: {line.strip()}")
                found = True
    except Exception as e:
        pass

if not found:
    print("Not found in any BAS or CLS file")
