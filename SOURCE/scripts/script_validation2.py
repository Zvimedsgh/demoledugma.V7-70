import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.017.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Range("rngPeriodValue").Validation.Delete', 'Range("rngPeriodValue").MergeArea.Validation.Delete')
content = content.replace('Range("rngPeriodValue").Validation.Add', 'Range("rngPeriodValue").MergeArea.Validation.Add')

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
