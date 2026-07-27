import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.188.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.189.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Add the reset of AA1 and AA2 in CreateInstructionSheets
old_create = """    ' Create or get Install Sheet
    Set wsInstall = Nothing"""

new_create = """    ' Reset the "Don't Show" flags so they appear for new users!
    wsMain.Range("AA1:AA2").ClearContents
    
    ' Create or get Install Sheet
    Set wsInstall = Nothing"""

content = content.replace(old_create, new_create)

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_188"', 'Attribute VB_Name = "Goren_Claude_V2_189"')
content = content.replace('VERSION: V2.188', 'VERSION: V2.189')
content = content.replace('APP_VERSION As String = "2.188"', 'APP_VERSION As String = "2.189"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.189 correctly!")
