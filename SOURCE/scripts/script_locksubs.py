import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code_1 = '''Public Sub BuildReview()

10      Dim wsMgmt As Worksheet'''

new_code_1 = '''Public Sub BuildReview()
    If CheckDemoLock() Then Exit Sub

10      Dim wsMgmt As Worksheet'''

old_code_2 = '''Public Sub SaveReportsToFolder()

10      On Error GoTo ERR_HANDLER'''

new_code_2 = '''Public Sub SaveReportsToFolder()
    If CheckDemoLock() Then Exit Sub

10      On Error GoTo ERR_HANDLER'''

if old_code_1 in content:
    content = content.replace(old_code_1, new_code_1)
if old_code_2 in content:
    content = content.replace(old_code_2, new_code_2)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
