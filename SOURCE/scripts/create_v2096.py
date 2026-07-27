import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.095.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.096.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_095"', 'Attribute VB_Name = "Goren_Claude_V2_096"')
content = content.replace("' VERSION: V2.095", "' VERSION: V2.096")
content = content.replace('Private Const APP_VERSION As String = "2.095"', 'Private Const APP_VERSION As String = "2.096"')

# Bypass already checked MsgBox
content = content.replace('If MsgBoxU(alreadyMsg, vbYesNo + vbQuestion) <> vbYes Then', 'If False Then')

# Bypass confirm MsgBox
content = content.replace('If MsgBoxU(confirmMsg, vbOKCancel + vbExclamation, CStr(wsMsgSrc.Cells(8, 1).Value)) <> vbOK Then', 'If False Then')

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.096 created.")
