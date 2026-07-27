import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.017.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

old_code1 = '''  70      wsMain.Range("rngPeriodValue").Validation.Delete'''
new_code1 = '''  70      wsMain.Range("rngPeriodValue").MergeArea.Validation.Delete'''

old_code2 = '''  200     wsMain.Range("rngPeriodValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listName'''
new_code2 = '''  200     wsMain.Range("rngPeriodValue").MergeArea.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listName'''

if old_code1 not in content or old_code2 not in content:
    print("Error: Old code not found in content")
    sys.exit(1)

content = content.replace(old_code1, new_code1)
content = content.replace(old_code2, new_code2)

# Also fix the empty validation when "shnatit" is chosen!
# Wait, actually it just does .Delete and then skips .Add.
# If we do .MergeArea.Validation.Delete it will completely remove the dropdown.

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Replacement successful")
