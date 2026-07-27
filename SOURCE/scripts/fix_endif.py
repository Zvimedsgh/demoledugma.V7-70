import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.055_20260702_1655.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if i+1 == 825 and line.strip() == "End If":
        pass # Skip the extra End If
    elif i+1 == 1455 and line.strip() == "End If":
        pass # Let's see if line 1455 has an extra End If too
    elif i+1 == 1490 and line.strip() == "End If":
        pass
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_LNUM.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)
