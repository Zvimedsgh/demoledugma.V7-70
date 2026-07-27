import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.145.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_title_color = """3450 wsMain.Range("A1").Font.Bold = True
If isDemoMode Then
wsMain.Range("A1").Font.Color = RGB(200, 0, 0) ' Red for Demo
Else
wsMain.Range("A1").Font.Color = RGB(0, 0, 0) ' Black for Real
End If"""

new_title_color = """3450 wsMain.Range("A1").Font.Bold = True
wsMain.Range("A1").Font.Color = RGB(200, 0, 0) ' Always Red!"""

if target_title_color in content:
    content = content.replace(target_title_color, new_title_color)
    print("Fixed title color logic.")
else:
    print("Could not find title color logic.")

# The user might be upset I changed the search title in V2.145 to red because they thought I hardcoded the AGENCY NAME there!
# "אתה לא יכול לשים אותה בקוד כי מחר אני ארצה לשים בהגרות שם של סוכנות אחרת"
# Wait, I didn't put the agency name in the search title, I just put "חיפוש לקוח חדש".
# I'll leave the search title alone for now. It's just red.

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_145"', 'Attribute VB_Name = "Goren_Claude_V2_146"')
content = content.replace('VERSION: V2.145', 'VERSION: V2.146')
content = content.replace('APP_VERSION As String = "2.145"', 'APP_VERSION As String = "2.146"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.146.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.146 created.")
