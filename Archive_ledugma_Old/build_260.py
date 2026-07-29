import os
import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.258.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.258"', 'APP_VERSION As String = "2.260"')
content = content.replace('VERSION: V2.258', 'VERSION: V2.260')
content = content.replace('Error in V2.258!', 'Error in V2.260!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_258"', 'Attribute VB_Name = "Goren_Claude_V2_260"')

# Update Changelog
changelog = """' CHANGES IN 2.260:
'   - BUGFIX: Fixed Error 1004 (We can't do that to a merged cell) by unmerging C16:L22 before applying credit.
"""
content = content.replace("' CHANGES IN 2.258:", changelog + "' CHANGES IN 2.258:")

# Fix the unmerge issue
# In V2.258, it was:
# 4791 On Error Resume Next
# 4792 wsMain.Range("C19:K20").UnMerge
# 4793 With wsMain.Range("C16:I16")
# I will change line 4792 to unmerge a large safe area to prevent any overlap issues.
content = content.replace('4792 wsMain.Range("C19:K20").UnMerge', '4792 wsMain.Range("A16:L22").UnMerge')

# Write V2.260
with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.260.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 260')
