import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1500.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_036"', 'Attribute VB_Name = "Goren_Claude_V2_040"')
content = content.replace('VERSION: V2.036', 'VERSION: V2.040')

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated Attribute VB_Name to V2_040")
