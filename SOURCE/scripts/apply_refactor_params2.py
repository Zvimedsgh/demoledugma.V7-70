import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_fixed.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(3850, 3890):
    lines[i] = lines[i].replace("wsMgmt", "wsParams")

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Replaced wsMgmt with wsParams in lines 3850-3890")
