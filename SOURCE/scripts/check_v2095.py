import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.095.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print("Application.Hwnd in file:", "Application.Hwnd" in content)

