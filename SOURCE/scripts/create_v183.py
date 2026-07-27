import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.182.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.183.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def split_long_vba_line(line, max_len=120):
    if len(line) <= max_len:
        return line
    if "ChrW" not in line:
        return line
        
    parts = line.split(" & ")
    new_line = ""
    current_line = ""
    
    for i, part in enumerate(parts):
        if i == 0:
            current_line = part
        else:
            if len(current_line) + len(part) + 3 > max_len:
                new_line += current_line + " & _\n        "
                current_line = part
            else:
                current_line += " & " + part
                
    new_line += current_line
    return new_line

for i in range(len(lines)):
    if len(lines[i]) > 200 and "ChrW(" in lines[i]:
        lines[i] = split_long_vba_line(lines[i], 120)

content = "".join(lines)
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_182"', 'Attribute VB_Name = "Goren_Claude_V2_183"')
content = content.replace('VERSION: V2.182', 'VERSION: V2.183')
content = content.replace('APP_VERSION As String = "2.182"', 'APP_VERSION As String = "2.183"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.183 correctly!")
