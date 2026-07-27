import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1700.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        lines[i] = re.sub(r'".*"', '"Goren_Claude_V2_056"', line)
    elif "VERSION: " in line:
        lines[i] = re.sub(r'VERSION: V\d\.\d+', 'VERSION: V2.056', line)
    elif "Private Const APP_VERSION As String =" in line:
        lines[i] = re.sub(r'".*"', '"2.056"', line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.056_20260702_1706")
