import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas', 'r', encoding='utf-8') as f:
    content = f.read()

old_sig = 'Private Sub BuildCompSlides(ByVal ppPres As Object, ByVal sheetName As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "", Optional ByVal excludeName As String = "", Optional ByVal incDocs As Boolean = True, Optional ByVal incInsured As Boolean = True)'
new_sig = 'Private Sub BuildCompSlides(ByVal ppPres As Object, ByVal sheetName As String, ByVal sheetTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "", Optional ByVal excludeName As String = "", Optional ByVal incDocs As Boolean = True, Optional ByVal incInsured As Boolean = True)'
content = content.replace(old_sig, new_sig)

old_titles = '''    Dim suffix As String
    suffix = ""
    If excludeName <> "" Then suffix = " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & excludeName
    tPrem = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & suffix
    tComm = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & suffix
    tDocs = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501) & suffix
    tIns = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501) & suffix'''

new_titles = '''    Dim suffix As String
    suffix = ""
    If excludeName <> "" Then suffix = " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & excludeName
    ' We append " lefi " (by) + sheetTitle + suffix
    Dim lefi As String
    lefi = " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " "
    tPrem = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & lefi & sheetTitle & suffix
    tComm = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & lefi & sheetTitle & suffix
    tDocs = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501) & lefi & sheetTitle & suffix
    tIns = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501) & lefi & sheetTitle & suffix'''

content = content.replace(old_titles, new_titles)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Updated BuildCompSlides to use sheetTitle')
