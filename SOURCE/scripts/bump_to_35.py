import sys
import shutil

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'

shutil.copy(filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("Goren_Claude_V2_034", "Goren_Claude_V2_035")
content = content.replace("V2.034", "V2.035")
content = content.replace('APP_VERSION As String = "2.034"', 'APP_VERSION As String = "2.035"')

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.035 successfully")
