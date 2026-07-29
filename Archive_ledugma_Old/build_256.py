import os
from datetime import datetime

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.255.bas', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('APP_VERSION As String = "2.255"', 'APP_VERSION As String = "2.256"')
content = content.replace('VERSION: V2.250', 'VERSION: V2.256')
content = content.replace('Error in V2.255!', 'Error in V2.256!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_255"', 'Attribute VB_Name = "Goren_Claude_V2_256"')
content = content.replace('DATE: 2026-07-08 17:44:11 (LATEST FIXES)', f'DATE: 2026-07-12 {datetime.now().strftime("%H:%M:%S")} (LATEST FIXES)')

# Let's insert a change log for 2.256
changelog = """' CHANGES IN 2.256:
'   - UI: Added credit text in D19, avoided clear conflicts, fixed cursor parking.
"""
insertion_point = content.find("' CHANGES IN 2.036:")
new_content = content[:insertion_point] + changelog + content[insertion_point:]

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.256.bas', 'w', encoding='utf-8') as f:
    f.write(new_content)
print('Generated 256')
