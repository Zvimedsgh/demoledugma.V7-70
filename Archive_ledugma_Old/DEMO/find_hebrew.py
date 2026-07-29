import glob

for filename in glob.glob(r"C:\ledugma\DEMO\modDemoReports*.bas"):
    with open(filename, "r", encoding="windows-1255", errors="ignore") as f:
        lines = f.readlines()
        for i, line in enumerate(lines):
            if "מבצע" in line or "מצגת" in line or "עיבוד" in line:
                print(f"File: {filename}, Line {i+1}: {line.strip()}")
