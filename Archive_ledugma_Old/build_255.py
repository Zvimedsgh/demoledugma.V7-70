import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.254.bas', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('APP_VERSION As String = "2.254"', 'APP_VERSION As String = "2.255"')
content = content.replace('VERSION: V2.254', 'VERSION: V2.255')
content = content.replace('Error in V2.254!', 'Error in V2.255!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_254"', 'Attribute VB_Name = "Goren_Claude_V2_255"')

# Replace the specific ClearContents in ApplyDemoLockOnOpen
# The easiest way is to split by lines, find ApplyDemoLockOnOpen, and replace it there
lines = content.split('\n')
in_apply_demo = False
for i, line in enumerate(lines):
    if 'Public Sub ApplyDemoLockOnOpen()' in line:
        in_apply_demo = True
    if 'End Sub' in line and in_apply_demo:
        in_apply_demo = False
    
    if in_apply_demo and 'wsMain.Range("D19:K20").ClearContents' in line:
        lines[i] = line.replace('D19:K20', 'J19:K20')

new_content = '\n'.join(lines)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.255.bas', 'w', encoding='utf-8') as f:
    f.write(new_content)
print('Generated 255')
