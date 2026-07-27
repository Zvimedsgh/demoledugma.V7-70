import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.210.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix tempPath in SendForReview
old_temp_path = """tempPath = REPORTS_FOLDER() & "\\" & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & "_" & wsRev.Name & ".xlsx\""""
new_temp_path = """tempPath = Environ$("TEMP") & "\\Exceptions_" & Format$(Now, "yyyy-mm-dd_hhmmss") & ".xlsx\""""
content = content.replace(old_temp_path, new_temp_path)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_210"', 'Attribute VB_Name = "Goren_Claude_V2_211"')
content = content.replace('VERSION: V2.210', 'VERSION: V2.211')
content = content.replace('APP_VERSION As String = "2.210"', 'APP_VERSION As String = "2.211"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.211")
