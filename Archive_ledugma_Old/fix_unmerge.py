import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.273.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Make the unmerge encompass the entire area that might be cleared
content = content.replace('330 wsMain.Range("A1:Z1").UnMerge', '330 wsMain.Range("A1:Z50").UnMerge')
# Keep ClearContents as it was
# 340 wsMain.Range("A1:Z1").ClearContents
# 350 wsMain.Range("A2:K20").ClearContents

# Also update the UnMerge lines that happen later to not fail if they are already unmerged.
# Actually, UnMerge on unmerged cells does nothing, it's safe.

content = content.replace('APP_VERSION As String = "2.273"', 'APP_VERSION As String = "2.274"')
content = content.replace('VERSION: V2.273', 'VERSION: V2.274')
content = content.replace('Error in V2.273!', 'Error in V2.274!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_273"', 'Attribute VB_Name = "Goren_Claude_V2_274"')

changelog = """' CHANGES IN 2.274:
'   - Bugfix: Setup error 1004 (We can't do that to a merged cell) at line 350 when resetting while in Laptop Mode.
"""
content = content.replace("' CHANGES IN 2.273:", changelog + "' CHANGES IN 2.273:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.274.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 274')
