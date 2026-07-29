import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.228.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.228', 'VERSION: V2.229')
content = content.replace('Error in V2.228!', 'Error in V2.229!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_228"', 'Attribute VB_Name = "Goren_Claude_V2_229"')
content = content.replace('APP_VERSION = "2.228"', 'APP_VERSION = "2.229"')

# Fix white cells in ApplyDemoLockOnOpen
clear_str = '    wsMain.Range("B14:K15").Clear\n    wsMain.Range("D19:K20").Clear'
new_clear_str = '    wsMain.Range("B13:K15").ClearContents\n    wsMain.Range("D18:K20").ClearContents\n    wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)\n    wsMain.Range("D18:K20").Interior.Color = RGB(220, 240, 220)'
content = content.replace(clear_str, new_clear_str)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.229.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 229')
