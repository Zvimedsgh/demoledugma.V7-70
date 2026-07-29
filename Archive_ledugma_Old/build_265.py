import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.264.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.264"', 'APP_VERSION As String = "2.265"')
content = content.replace('VERSION: V2.264', 'VERSION: V2.265')
content = content.replace('Error in V2.264!', 'Error in V2.265!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_264"', 'Attribute VB_Name = "Goren_Claude_V2_265"')

# Changelog
changelog = """' CHANGES IN 2.265:
'   - BUGFIX: Fully addressed the root cause of the "Line 1742" crash which was actually the "Clear previous mess" block intersecting with V2.259's leftover merged cells.
"""
content = content.replace("' CHANGES IN 2.264:", changelog + "' CHANGES IN 2.264:")

old_clear_block = '''    ' Clear previous mess
    wsMain.Range("B13:K15").ClearContents
    wsMain.Range("D18:K20").ClearContents
    wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)
    wsMain.Range("D18:K20").Interior.Color = RGB(220, 240, 220)
    On Error Resume Next'''

new_clear_block = '''    ' Clear previous mess
    On Error Resume Next
    wsMain.Range("A12:L20").UnMerge
    wsMain.Range("A12:L20").ClearContents
    wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)
    wsMain.Range("D18:K20").Interior.Color = RGB(220, 240, 220)'''

content = content.replace(old_clear_block, new_clear_block)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.265.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 265')
