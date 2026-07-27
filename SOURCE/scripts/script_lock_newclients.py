import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''Public Sub NewClients()
    Dim wsMain As Worksheet'''

new_code = '''Public Sub NewClients()
    If CheckDemoLock() Then Exit Sub
    Dim wsMain As Worksheet'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
