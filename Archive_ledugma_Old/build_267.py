import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.266.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.266"', 'APP_VERSION As String = "2.267"')
content = content.replace('VERSION: V2.266', 'VERSION: V2.267')
content = content.replace('Error in V2.266!', 'Error in V2.267!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_266"', 'Attribute VB_Name = "Goren_Claude_V2_267"')

# Replace the toggle macro with a refined one
old_macro = """        Call get_rates
        
    Else
        ' Move from Laptop to Desktop
        wsMain.Range("F12:G14").ClearContents
        wsMain.Range("F12:G14").Interior.Color = xlNone
        wsMain.Range("F12:G14").Borders.LineStyle = xlNone"""

new_macro = """        Call get_rates
        
    Else
        ' Move from Laptop to Desktop
        wsMain.Range("F12:G14").ClearContents
        wsMain.Range("F12:G14").Interior.Color = xlNone
        wsMain.Range("F12:G14").Borders.LineStyle = xlNone
        
        ' Restore background color for demo message if needed
        If wsMain.Range("B13").Interior.Color = RGB(220, 240, 220) Then
            wsMain.Range("B13:K15").Interior.Color = RGB(220, 240, 220)
        End If"""

content = content.replace(old_macro, new_macro)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.267.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 267')
