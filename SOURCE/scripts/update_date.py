import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.101.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("DATE: 2026-07-02 12:57", "DATE: 2026-07-03")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.101 date updated.")
