import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.262.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.262"', 'APP_VERSION As String = "2.263"')
content = content.replace('VERSION: V2.262', 'VERSION: V2.263')
content = content.replace('Error in V2.262!', 'Error in V2.263!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_262"', 'Attribute VB_Name = "Goren_Claude_V2_263"')

# Changelog
changelog = """' CHANGES IN 2.263:
'   - BUGFIX: Forced On Error Resume Next across the entire G3/G4 validation cleanup block.
"""
content = content.replace("' CHANGES IN 2.262:", changelog + "' CHANGES IN 2.262:")

# Replace the specific On Error GoTo ERR_HANDLER before 1760
old_err_block = '''    wsMain.Shapes("shpDemoMsgText").Delete
    wsMain.Range("D19:K20").ClearContents
    On Error GoTo ERR_HANDLER
    1760 With wsMain.Range("G3").MergeArea.Validation'''

new_err_block = '''    wsMain.Shapes("shpDemoMsgText").Delete
    wsMain.Range("D19:K20").ClearContents
    On Error Resume Next
    1760 With wsMain.Range("G3").MergeArea.Validation'''

content = content.replace(old_err_block, new_err_block)

# Just to be 100% certain, clear Err object at the end of the block so it doesn't leak
old_end_block = '''        1800 .InCellDropdown = True
    1810 End With
        1820 On Error Resume Next'''

new_end_block = '''        1800 .InCellDropdown = True
    1810 End With
    Err.Clear
    1820 On Error GoTo ERR_HANDLER'''
content = content.replace(old_end_block, new_end_block)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.263.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 263')
