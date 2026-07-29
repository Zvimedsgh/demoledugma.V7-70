import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.01.bas', 'r', encoding='utf-8') as f:
    content = f.read()

msgbox_old = """    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
    
    ' Bring PowerPoint to the front after user clicks OK
    On Error Resume Next
    AppActivate "PowerPoint"
    On Error GoTo ERR_HANDLER"""

msgbox_new = """    ' Bring PowerPoint to the front first!
    On Error Resume Next
    AppActivate "PowerPoint"
    On Error GoTo ERR_HANDLER
    
    ' Show message box ON TOP of PowerPoint using vbSystemModal (4096)
    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation + vbSystemModal"""

content = content.replace(msgbox_old, msgbox_new)

content = content.replace('APP_VERSION As String = "3.01"', 'APP_VERSION As String = "3.02"')
content = content.replace('VERSION: V3.01', 'VERSION: V3.02')
content = content.replace('Error in V3.01!', 'Error in V3.02!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_01"', 'Attribute VB_Name = "Goren_Claude_V3_02"')

changelog = """' CHANGES IN 3.02:
'   - UI: Swapped the order so PowerPoint is brought to front BEFORE the success message, and made the message TopMost (SystemModal) so it appears clearly over PowerPoint.
"""
content = content.replace("' CHANGES IN 3.01:", changelog + "' CHANGES IN 3.01:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.02.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.02')
