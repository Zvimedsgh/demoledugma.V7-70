import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.026.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.027.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_026", "Goren_Claude_V2_027")
content = content.replace("V2.026", "V2.027")
content = content.replace('APP_VERSION As String = "2.026"', 'APP_VERSION As String = "2.027"')

# Unlock Button 1
old_lock = """Public Sub BuildReview()
    If CheckDemoLock() Then Exit Sub"""

new_lock = """Public Sub BuildReview()
    ' Unlocked in Demo Mode so it can process the fixed DATA sheets"""

content = content.replace(old_lock, new_lock)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.027 successfully")
