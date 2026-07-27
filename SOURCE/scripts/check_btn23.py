import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.068.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
curr_func = ""
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub ApplyCorrectionsAndBuildReports()" in line_strip:
        in_func = True
        curr_func = "ApplyCorrections"
    elif "Public Sub BuildPresentation()" in line_strip:
        in_func = True
        curr_func = "BuildPresentation"
    elif in_func and "End Sub" in line_strip:
        in_func = False
    elif in_func and "Exit Sub" in line_strip:
        print(f"[{i+1}] {curr_func}: {line_strip}")

