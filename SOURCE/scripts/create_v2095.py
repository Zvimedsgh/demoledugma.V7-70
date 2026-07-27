import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.095.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_094"', 'Attribute VB_Name = "Goren_Claude_V2_095"')
content = content.replace("' VERSION: V2.094", "' VERSION: V2.095")
content = content.replace('Private Const APP_VERSION As String = "2.094"', 'Private Const APP_VERSION As String = "2.095"')

# Replace MsgBoxU definition using regex
new_msgboxu = """Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
    Dim hwnd As LongPtr
    On Error Resume Next
    hwnd = Application.Hwnd
    On Error GoTo 0
    MsgBoxU = MessageBoxW(hwnd, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT)
End Function"""

pattern = re.compile(r'Private Function MsgBoxU.*?End Function', re.DOTALL)
content = pattern.sub(new_msgboxu, content)

with open(outpath, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.095 created with Regex.")
