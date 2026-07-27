import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.120.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

content = "".join(lines)

start_str = "Dim askMsg As String"
idx = content.rfind(start_str)
print(content[idx:idx+1500])

