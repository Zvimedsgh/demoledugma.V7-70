import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.163.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.164.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_163"', 'Attribute VB_Name = "Goren_Claude_V2_164"')
content = content.replace('VERSION: V2.163', 'VERSION: V2.164')
content = content.replace('APP_VERSION As String = "2.163"', 'APP_VERSION As String = "2.164"')

# Fix the duplicate declaration
bad_block = """    Dim tLeft As Single, tWidth As Single
    tLeft = wsMain.Range("F19").Left
    tWidth = wsMain.Range("I19").Left + wsMain.Range("I19").Width - tLeft
    Dim tLeft As Single, tWidth As Single
    tLeft = wsMain.Range("E19").Left
    tWidth = wsMain.Range("K19").Left + wsMain.Range("K19").Width - tLeft
    Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, tLeft, wsMain.Range("F19").Top, tWidth, 40)"""

good_block = """    Dim tLeft As Single, tWidth As Single
    tLeft = wsMain.Range("E19").Left
    tWidth = wsMain.Range("K19").Left + wsMain.Range("K19").Width - tLeft
    Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, tLeft, wsMain.Range("F19").Top, tWidth, 40)"""

if bad_block in content:
    content = content.replace(bad_block, good_block)
else:
    # Just to be safe if indentation doesn't perfectly match, let's use regex
    pattern = r'Dim tLeft As Single, tWidth As Single.*?Dim tLeft As Single, tWidth As Single'
    content = re.sub(pattern, 'Dim tLeft As Single, tWidth As Single', content, flags=re.DOTALL)

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.164 correctly!")
