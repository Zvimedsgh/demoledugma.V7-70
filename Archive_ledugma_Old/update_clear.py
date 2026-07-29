import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.275.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Expand clear range to wipe out old lingering text from previous versions
content = content.replace('350 wsMain.Range("A2:K20").ClearContents', '350 wsMain.Range("A2:L50").ClearContents')

content = content.replace('APP_VERSION As String = "2.275"', 'APP_VERSION As String = "2.276"')
content = content.replace('VERSION: V2.275', 'VERSION: V2.276')
content = content.replace('Error in V2.275!', 'Error in V2.276!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_275"', 'Attribute VB_Name = "Goren_Claude_V2_276"')

changelog = """' CHANGES IN 2.276:
'   - UI: Expanded screen clear area to A2:L50 to remove leftover artifacts (like old credit text in row 24).
"""
content = content.replace("' CHANGES IN 2.275:", changelog + "' CHANGES IN 2.275:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.276.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 276')
