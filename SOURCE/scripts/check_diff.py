import sys

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.198.bas', 'r', encoding='utf-8') as f:
    lines198 = f.readlines()

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas', 'r', encoding='utf-8') as f:
    lines199 = f.readlines()

print("V2.198 line 301:", lines198[300].strip())
print("V2.199 line 301:", lines199[300].strip())

