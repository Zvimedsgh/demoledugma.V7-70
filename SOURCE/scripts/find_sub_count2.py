import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

count1 = 0
count2 = 0
for i, line in enumerate(lines):
    if "Public Sub UpdatePeriodDropdown" in line:
        count1 += 1
    if "Public Sub UpdateFilterValueDropdown" in line:
        count2 += 1
print(f"UpdatePeriodDropdown: {count1}")
print(f"UpdateFilterValueDropdown: {count2}")
