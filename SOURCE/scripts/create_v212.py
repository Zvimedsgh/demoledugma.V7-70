import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.211.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.212.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Remove the Minimize line and Min/Max lines, replace with just Maximize and add MsgBoxU
pattern = r'(If Not ppApp Is Nothing And Not ppPres Is Nothing Then\s*)Application\.WindowState = xlMinimized\s*ppApp\.Visible = True\s*ppApp\.WindowState = 2 \' Min\s*ppApp\.WindowState = 3 \' Max\s*ppApp\.Activate\s*End If'
replacement = r'\1ppApp.Visible = True\n        ppApp.WindowState = 3 \' Max\n        ppApp.Activate\n        End If'
content = re.sub(pattern, replacement, content)

# Add MsgBoxU right before 1060 Exit Sub
pattern2 = r'(Application\.DisplayAlerts = True\s*)1060\s*Exit Sub'
replacement2 = r'\1\' Show success message\nMsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation\n\n1060    Exit Sub'
content = re.sub(pattern2, replacement2, content)


# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_211"', 'Attribute VB_Name = "Goren_Claude_V2_212"')
content = content.replace('VERSION: V2.211', 'VERSION: V2.212')
content = content.replace('APP_VERSION As String = "2.211"', 'APP_VERSION As String = "2.212"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.212")
