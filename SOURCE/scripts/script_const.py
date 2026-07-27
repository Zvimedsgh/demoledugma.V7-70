import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''Private Const PARAM_PREMIUM_THRESHOLD As String = "PREMIUM_THRESHOLD"'''
new_code = '''Private Const PARAM_PREMIUM_THRESHOLD As String = "PREMIUM_THRESHOLD"
Private Const PARAM_AGENCY_NAME As String = "AGENCY_NAME"
Private Const PARAM_DEMO_AGENCY_NAME As String = "DEMO_AGENCY_NAME"'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
