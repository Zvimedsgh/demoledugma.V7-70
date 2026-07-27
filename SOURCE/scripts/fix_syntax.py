import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.049.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    new_lines.append(line.replace("Then isDemoMode = True Then isDemoMode = True", "Then isDemoMode = True"))

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.049_20260702_1635.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.049 fixed")
