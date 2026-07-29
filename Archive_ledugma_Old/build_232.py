import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.231.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.231', 'VERSION: V2.232')
content = content.replace('Error in V2.231!', 'Error in V2.232!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_231"', 'Attribute VB_Name = "Goren_Claude_V2_232"')
content = content.replace('APP_VERSION As String = "2.231"', 'APP_VERSION As String = "2.232"')
content = content.replace('APP_VERSION = "2.231"', 'APP_VERSION = "2.232"')

# Move shpManualMsg to L1
content = content.replace('wsMain.Range("L2").Left, wsMain.Range("L2").Top', 'wsMain.Range("L1").Left, wsMain.Range("L1").Top')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.232.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 232')
