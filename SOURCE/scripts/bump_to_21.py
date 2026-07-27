import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.020.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.021.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("Goren_Claude_V2_020", "Goren_Claude_V2_021")
content = content.replace("V2.020", "V2.021")
content = content.replace('APP_VERSION As String = "2.020"', 'APP_VERSION As String = "2.021"')

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.021 successfully")
