import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.166.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.167.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_166"', 'Attribute VB_Name = "Goren_Claude_V2_167"')
content = content.replace('VERSION: V2.166', 'VERSION: V2.167')
content = content.replace('APP_VERSION As String = "2.166"', 'APP_VERSION As String = "2.167"')

# Fix the merged cell error by removing wsMain.Range("F19").ClearContents
old_line = '3970 wsMain.Range("F19").ClearContents'
if old_line in content:
    content = content.replace(old_line, "' " + old_line)
else:
    print("Could not find the line!")

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.167 correctly!")
