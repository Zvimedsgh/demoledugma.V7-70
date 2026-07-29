import re
with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# find any paths that have C:\
for line in text.split('\n'):
    if 'C:\\' in line.upper():
        print(line.strip())
