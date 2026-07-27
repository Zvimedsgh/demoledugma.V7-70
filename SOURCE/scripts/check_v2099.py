import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print("Application.Hwnd:", "Application.Hwnd" in content)
print("WinHttpRequest:", "WinHttpRequest.5.1" in content)
print("srcData = wsSrc.Range:", "srcData = wsSrc.Range" in content)

