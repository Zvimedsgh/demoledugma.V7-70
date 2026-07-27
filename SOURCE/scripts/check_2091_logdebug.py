import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.091.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print("Private Sub LogDebug in file:", "Private Sub LogDebug" in content)
print("Option Explicit in file:", "Option Explicit" in content)

