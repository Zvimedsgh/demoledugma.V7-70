import re

with open(r'C:\ledugma\DEMO\modDemoReports_V9.36.bas', encoding='windows-1255', errors='ignore') as f:
    text = f.read()

ranges = set(re.findall(r'wsMain\.Range\("([A-Z][0-9]+(?::[A-Z][0-9]+)?)"\)', text))
print("wsMain.Range addresses:", sorted(list(ranges)))

ranges_this = set(re.findall(r'ThisWorkbook\.Worksheets\(CONTROL_SHEET_NAME\(\)\)\.Range\("([A-Z][0-9]+(?::[A-Z][0-9]+)?)"\)', text))
print("ThisWorkbook...Range addresses:", sorted(list(ranges_this)))

names = set(re.findall(r'RefersTo:="=\'" & wsMain\.Name & "\'!([\$A-Z0-9:]+)"', text))
print("Named ranges RefersTo:", sorted(list(names)))
