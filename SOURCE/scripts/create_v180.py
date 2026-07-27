import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.179.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.180.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(len(lines)):
    if 'shp.TextFrame2.MarginRight = 10' in lines[i]:
        # The line has a number like '2950' or '2955'.
        # Let's replace whatever number is there with a number ending in 5 based on the previous line.
        prev_line = lines[i-1]
        m = re.match(r'\s*(\d+)', prev_line)
        if m:
            new_num = int(m.group(1)) + 5
            # Replace the leading number on the current line
            lines[i] = re.sub(r'^\s*(\d+)', f'    {new_num}', lines[i])
        else:
            lines[i] = re.sub(r'^\s*(\d+)', '', lines[i])

content = "".join(lines)
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_179"', 'Attribute VB_Name = "Goren_Claude_V2_180"')
content = content.replace('VERSION: V2.179', 'VERSION: V2.180')
content = content.replace('APP_VERSION As String = "2.179"', 'APP_VERSION As String = "2.180"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.180 correctly!")
