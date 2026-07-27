import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.194.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.195.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix HideWorkSheets logic for AA1
old_hide = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' \u05d4_\u05d4
If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then hideIt = False
End If"""

new_hide = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then
hideIt = False
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
End If
End If"""
content = content.replace(old_hide, new_hide)

# Fix CheckUserPermissions logic for Limited Access
old_chk = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' \u05d4_\u05d4
If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then hideIt2 = False
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' \u05d4_\u05ea
If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then hideIt2 = False
End If"""

new_chk = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then
hideIt2 = False
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
End If
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then
hideIt2 = False
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
End If
End If"""
content = content.replace(old_chk, new_chk)


# Update version strings
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_194"', 'Attribute VB_Name = "Goren_Claude_V2_195"')
content = content.replace('VERSION: V2.194', 'VERSION: V2.195')
content = content.replace('APP_VERSION As String = "2.194"', 'APP_VERSION As String = "2.195"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.195")
