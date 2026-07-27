import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.049_20260702_1635.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if 'yearVal = "2025"' in line and "isDemoMode" in lines[i-1]:
        # Delete this line by skipping
        pass
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.050 with fixed yearVal")
