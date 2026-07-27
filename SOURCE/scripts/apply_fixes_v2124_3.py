import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if 'wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)' in line and '650' in str(i):
        new_lines.append(line.replace('ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)', '""'))
    elif 'wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)' in line and '650' in str(i):
        new_lines.append(line)
        new_lines.append('On Error Resume Next\n')
        new_lines.append('wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)\n')
        new_lines.append('On Error GoTo 0\n')
    else:
        new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Replaced robustly.")
