import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.134.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_search = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If UCase(cName) = UCase(searchText) Then bMatch = True
End If"""

new_search = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If InStr(1, cName, searchText, vbTextCompare) > 0 Then bMatch = True
End If"""

if target_search in content:
    content = content.replace(target_search, new_search)
    print("Search logic updated.")
else:
    print("Could not find search logic.")

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

