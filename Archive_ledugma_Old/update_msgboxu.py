import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.06.bas', 'r', encoding='utf-8') as f:
    content = f.read()

msgbox_old = """Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
Dim hwnd As LongPtr
On Error Resume Next
hwnd = Application.Hwnd
On Error GoTo 0
MsgBoxU = MessageBoxW(hwnd, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT)
End Function"""

msgbox_new = """Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
Dim hwnd As LongPtr
' If SystemModal is requested, detach from Excel (hwnd=0) so it floats above all apps
If (uType And vbSystemModal) = vbSystemModal Then
    hwnd = 0
    uType = uType Or vbMsgBoxSetForeground
Else
    On Error Resume Next
    hwnd = Application.Hwnd
    On Error GoTo 0
End If
MsgBoxU = MessageBoxW(hwnd, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT)
End Function"""

content = content.replace(msgbox_old, msgbox_new)

content = content.replace('APP_VERSION As String = "3.06"', 'APP_VERSION As String = "3.07"')
content = content.replace('VERSION: V3.06', 'VERSION: V3.07')
content = content.replace('Error in V3.06!', 'Error in V3.07!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_06"', 'Attribute VB_Name = "Goren_Claude_V3_07"')

changelog = """' CHANGES IN 3.07:
'   - BUGFIX: Made MsgBoxU detach from Excel's window handle if vbSystemModal is used, ensuring the "Finished" message pops up ON TOP of PowerPoint rather than staying grouped with Excel.
"""
content = content.replace("' CHANGES IN 3.06:", changelog + "' CHANGES IN 3.06:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.07.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.07')
