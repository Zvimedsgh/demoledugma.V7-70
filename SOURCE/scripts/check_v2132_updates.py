import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.132.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

found_expl = False
found_credit = False

for i, line in enumerate(lines):
    if "wsSearch.Range(\"E5:E8\").Font.Size = 12" in line:
        found_expl = True
    if "wsMain.Range(\"A22\").Value = ChrW" in line:
        found_credit = True

print(f"Explanation updated: {found_expl}")
print(f"Credit updated: {found_credit}")

