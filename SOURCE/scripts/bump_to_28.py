import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.027.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_027", "Goren_Claude_V2_028")
content = content.replace("V2.027", "V2.028")
content = content.replace('APP_VERSION As String = "2.027"', 'APP_VERSION As String = "2.028"')

# Fix MIN_PREMIUM value in A00_SetupMainSheet
old_val = 'wsParams.Cells(2, 2).Value = "20"'
new_val = 'wsParams.Cells(2, 2).Value = "200000"'
content = content.replace(old_val, new_val)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.028 successfully")
