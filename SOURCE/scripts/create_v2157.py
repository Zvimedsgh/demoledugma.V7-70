import sys
import shutil

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.156.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.157.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_156"', 'Attribute VB_Name = "Goren_Claude_V2_157"')
content = content.replace('VERSION: V2.156', 'VERSION: V2.157')
content = content.replace('APP_VERSION As String = "2.156"', 'APP_VERSION As String = "2.157"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.157 successfully!")
