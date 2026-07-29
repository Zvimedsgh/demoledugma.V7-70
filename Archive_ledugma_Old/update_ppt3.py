import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.07.bas', 'r', encoding='utf-8') as f:
    content = f.read()

target_old = """' Automatically open PowerPoint without prompt (as requested by user)
On Error Resume Next
If Not ppApp Is Nothing And Not ppPres Is Nothing Then
ppApp.Visible = True
ppApp.WindowState = 3 ' Max
ppApp.Activate
End If
On Error GoTo ERR_HANDLER
910     Set ppPres = Nothing
920     Set ppApp = Nothing


' Clean up presentation sheets automatically
Application.DisplayAlerts = False
' (REMOVED: User requested that result sheets are NOT deleted after presentation)
Application.DisplayAlerts = True

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True
' Show success message
MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation

1060    Exit Sub"""

target_new = """' Clean up presentation sheets automatically
Application.DisplayAlerts = False
' (REMOVED: User requested that result sheets are NOT deleted after presentation)
Application.DisplayAlerts = True

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.DisplayAlerts = True

' Show success message in Excel FIRST
MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & vbCrLf & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512), vbInformation

' Automatically open PowerPoint without prompt AFTER the user clicks OK
On Error Resume Next
If Not ppApp Is Nothing And Not ppPres Is Nothing Then
    ppApp.Visible = True
    ppApp.WindowState = 3 ' Max
    ppApp.Activate
    AppActivate "PowerPoint"
End If
On Error GoTo ERR_HANDLER

910     Set ppPres = Nothing
920     Set ppApp = Nothing

1060    Exit Sub"""

content = content.replace(target_old, target_new)

# Also remove the vbSystemModal hack from MsgBoxU since we don't need it anymore if we show it in Excel first
msgbox_old = """Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
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

msgbox_new = """Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
Dim hwnd As LongPtr
On Error Resume Next
hwnd = Application.Hwnd
On Error GoTo 0
MsgBoxU = MessageBoxW(hwnd, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT)
End Function"""

content = content.replace(msgbox_old, msgbox_new)


content = content.replace('APP_VERSION As String = "3.07"', 'APP_VERSION As String = "3.08"')
content = content.replace('VERSION: V3.07', 'VERSION: V3.08')
content = content.replace('Error in V3.07!', 'Error in V3.08!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_07"', 'Attribute VB_Name = "Goren_Claude_V3_08"')

changelog = """' CHANGES IN 3.08:
'   - UI: Reordered presentation completion so the message box appears natively in Excel FIRST, explicitly telling the user to click OK to open PowerPoint. Once clicked, PowerPoint maximizes and jumps to the front.
"""
content = content.replace("' CHANGES IN 3.07:", changelog + "' CHANGES IN 3.07:")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.08.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 3.08')
