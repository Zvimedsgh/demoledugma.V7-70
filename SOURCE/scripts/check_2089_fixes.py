import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.089.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

print("Static apiFailed in file:", "Static apiFailed" in content)
print("apiCallsCount in file:", "apiCallsCount" in content)
print("matchBr = Application.Match in file:", "matchBr = Application.Match" in content)
print("yr >= 2000 And yr <= 2050 in file:", "yr >= 2000" in content)
print("dateArr = wsSource.Range" in content)

