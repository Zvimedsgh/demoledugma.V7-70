import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.214.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Dim wsName1 As String, wsName2 As String" in line:
        print("Replacement SUCCESS")
        break
else:
    print("Replacement FAILED")

