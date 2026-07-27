import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.051_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    new_lines.append(line)
    if "If sName = H_SET_PARAMS() Then IsProtectedSheet = True: Exit Function" in line:
        new_lines.append("    If UCase$(Left$(sName, 5)) = \"DATA_\" Then IsProtectedSheet = True: Exit Function\n")

# Let's also bump version
for i, line in enumerate(new_lines):
    if "Attribute VB_Name =" in line:
        new_lines[i] = line.replace("V2_051", "V2_052")
    if "Private Const APP_VERSION As String =" in line:
        new_lines[i] = line.replace("2.051", "2.052")
    if "VERSION: V2.051" in line:
        new_lines[i] = line.replace("2.051", "2.052")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.052.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.052")
