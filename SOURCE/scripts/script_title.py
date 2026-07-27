import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code = '''  100     shp.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492)'''
new_code = '''  100     shp.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492) & " - " & GetActiveAgencyName()'''

if old_code not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code, new_code)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
