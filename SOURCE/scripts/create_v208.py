import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.207.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.208.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix SOURCE_FOLDER
old_sf = """If Mid$(SOURCE_FOLDER, 2, 2) = ":\\\\" Then SOURCE_FOLDER = Left$(SOURCE_FOLDER, 2) & "\\\\" & Mid$(SOURCE_FOLDER, 4)"""
new_sf = """SOURCE_FOLDER = Replace(SOURCE_FOLDER, ":\\\\", ":\\")"""
content = content.replace(old_sf, new_sf)

# Fix REPORTS_FOLDER
old_rf = """If Mid$(REPORTS_FOLDER, 2, 2) = ":\\\\" Then REPORTS_FOLDER = Left$(REPORTS_FOLDER, 2) & "\\\\" & Mid$(REPORTS_FOLDER, 4)"""
new_rf = """REPORTS_FOLDER = Replace(REPORTS_FOLDER, ":\\\\", ":\\")"""
content = content.replace(old_rf, new_rf)

# Fix BACKUP_FOLDER
old_bf = """If Mid$(BACKUP_FOLDER, 2, 2) = ":\\\\" Then BACKUP_FOLDER = Left$(BACKUP_FOLDER, 2) & "\\\\" & Mid$(BACKUP_FOLDER, 4)"""
new_bf = """BACKUP_FOLDER = Replace(BACKUP_FOLDER, ":\\\\", ":\\")"""
content = content.replace(old_bf, new_bf)


# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_207"', 'Attribute VB_Name = "Goren_Claude_V2_208"')
content = content.replace('VERSION: V2.207', 'VERSION: V2.208')
content = content.replace('APP_VERSION As String = "2.207"', 'APP_VERSION As String = "2.208"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.208")
