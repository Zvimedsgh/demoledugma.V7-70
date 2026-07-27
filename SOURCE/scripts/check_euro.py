import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print("3.4239 in file:", "3.4239" in content)

