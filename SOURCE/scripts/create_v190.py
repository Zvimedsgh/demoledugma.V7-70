import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.189.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.190.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update HideWorkSheets
old_hide1 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then hideIt = False
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then hideIt = False
End If"""

new_hide1 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
    If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA1").Value) <> "YES" Then 
        hideIt = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
    If UCase$(ThisWorkbook.Worksheets(ctrlName).Range("AA2").Value) <> "YES" Then 
        hideIt = False
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    End If
End If"""

content = content.replace(old_hide1, new_hide1)

# 2. Update CheckUserPermissions
old_hide2 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA1").Value) <> "YES" Then hideIt2 = False
End If
If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1514) & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) Then ' הוראות_תפעול
If UCase$(ThisWorkbook.Worksheets(homeSheetName).Range("AA2").Value) <> "YES" Then hideIt2 = False
End If"""

new_hide2 = """If ws.Name = ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1494) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) Then ' הוראות_התקנה
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

content = content.replace(old_hide2, new_hide2)


content = content.replace('Attribute VB_Name = "Goren_Claude_V2_189"', 'Attribute VB_Name = "Goren_Claude_V2_190"')
content = content.replace('VERSION: V2.189', 'VERSION: V2.190')
content = content.replace('APP_VERSION As String = "2.189"', 'APP_VERSION As String = "2.190"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.190 correctly!")
