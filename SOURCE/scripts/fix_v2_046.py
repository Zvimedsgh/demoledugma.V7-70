import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_next = 0
for i, line in enumerate(lines):
    if skip_next > 0:
        skip_next -= 1
        continue
    
    if "Dim tHalf As String, tQuarter As String, tMonth As String" in line:
        # We delete this and the 3 lines after it
        skip_next = 3
        continue
        
    if "' --- Build Maps for Native Dropdowns ---" in line:
        new_lines.append('Dim tHalf As String, tQuarter As String, tMonth As String\n')
        new_lines.append('tHalf = ChrW(1495) & ChrW(1510) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)\n')
        new_lines.append('tQuarter = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503)\n')
        new_lines.append('tMonth = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497)\n')
        new_lines.append(line)
        continue

    new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.046.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.046 with variable declaration fix")
