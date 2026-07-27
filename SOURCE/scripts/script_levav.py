import sys
filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.019.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = 'ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489)'

if target in content:
    content = content.replace(target, 'GetActiveAgencyName()')
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Replaced all occurrences of Levav System.")
else:
    print("Not found.")
