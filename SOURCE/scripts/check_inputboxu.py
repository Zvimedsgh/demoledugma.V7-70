import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.111.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

if "Function InputBoxU" in content:
    print("InputBoxU exists")
else:
    print("InputBoxU DOES NOT EXIST")
