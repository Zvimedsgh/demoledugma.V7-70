import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print("3.002 in file:", "3.002" in content)
print("3.4239 in file:", "3.4239" in content)
print("WinHttpRequest in file:", "WinHttpRequest.5.1" in content)
print("SetTimeouts 2000, 2000 in file:", "2000, 2000, 2000, 2000" in content)
