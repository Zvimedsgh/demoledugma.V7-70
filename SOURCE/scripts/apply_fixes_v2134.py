import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.133.bas'

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

# Update Whatsapp
# Let's find exactly the text set in Demo msg
wa_text_target = """shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp  " & ChrW(1500) & " - 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 17"""

wa_text_new = """shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp" & ChrW(1500) & " - 054-6677396"
shpDemoMsg.TextFrame2.TextRange.Font.Size = 17
shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter"""

if wa_text_target in content:
    content = content.replace(wa_text_target, wa_text_new)
    print("Whatsapp text updated.")
else:
    print("Could not find whatsapp text.")

# Make sure we hit the second instance of whatsapp text as well (ApplyDemoLockOnOpen)
# Wait, my target includes shpDemoMsg.TextFrame2.TextRange.Font.Size = 17, which should exist in both.
# Let's see if the replace did both.

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_133"', 'Attribute VB_Name = "Goren_Claude_V2_134"')
content = content.replace('VERSION: V2.133', 'VERSION: V2.134')
content = content.replace('APP_VERSION As String = "2.133"', 'APP_VERSION As String = "2.134"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.134.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.134 created.")
