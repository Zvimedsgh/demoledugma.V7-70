import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.206.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.207.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix SOURCE_FOLDER
old_sf = """If Right$(SOURCE_FOLDER, 1) <> "\\" Then SOURCE_FOLDER = SOURCE_FOLDER & "\\"
End Function"""
new_sf = """If Right$(SOURCE_FOLDER, 1) <> "\\" Then SOURCE_FOLDER = SOURCE_FOLDER & "\\"
If Mid$(SOURCE_FOLDER, 2, 2) = ":\\\\" Then SOURCE_FOLDER = Left$(SOURCE_FOLDER, 2) & "\\" & Mid$(SOURCE_FOLDER, 4)
End Function"""
content = content.replace(old_sf, new_sf)

# Fix REPORTS_FOLDER
old_rf = """If Right$(REPORTS_FOLDER, 1) <> "\\" Then REPORTS_FOLDER = REPORTS_FOLDER & "\\"
End Function"""
new_rf = """If Right$(REPORTS_FOLDER, 1) <> "\\" Then REPORTS_FOLDER = REPORTS_FOLDER & "\\"
If Mid$(REPORTS_FOLDER, 2, 2) = ":\\\\" Then REPORTS_FOLDER = Left$(REPORTS_FOLDER, 2) & "\\" & Mid$(REPORTS_FOLDER, 4)
End Function"""
content = content.replace(old_rf, new_rf)

# Fix BACKUP_FOLDER
old_bf = """If Right$(BACKUP_FOLDER, 1) <> "\\" Then BACKUP_FOLDER = BACKUP_FOLDER & "\\"
End Function"""
new_bf = """If Right$(BACKUP_FOLDER, 1) <> "\\" Then BACKUP_FOLDER = BACKUP_FOLDER & "\\"
If Mid$(BACKUP_FOLDER, 2, 2) = ":\\\\" Then BACKUP_FOLDER = Left$(BACKUP_FOLDER, 2) & "\\" & Mid$(BACKUP_FOLDER, 4)
End Function"""
content = content.replace(old_bf, new_bf)


# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_206"', 'Attribute VB_Name = "Goren_Claude_V2_207"')
content = content.replace('VERSION: V2.206', 'VERSION: V2.207')
content = content.replace('APP_VERSION As String = "2.206"', 'APP_VERSION As String = "2.207"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.207")
