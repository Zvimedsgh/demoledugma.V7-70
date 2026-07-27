import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.120.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# --- Task 1: Add DisplayAlerts=False to DeleteSheetIfExists ---
target_del = """' Delete the sheet
80      On Error Resume Next
90      wsTarget.Delete
100     On Error GoTo 0"""

new_del = """' Delete the sheet
80      On Error Resume Next
        Application.DisplayAlerts = False
90      wsTarget.Delete
        Application.DisplayAlerts = True
100     On Error GoTo 0"""

if target_del in content:
    content = content.replace(target_del, new_del)
else:
    print("Could not find target_del")

# --- Task 2: Remove automatic slideshow running ---
target_slide = """On Error GoTo ERR_HANDLER        ' Start slideshow automatically
On Error Resume Next
If Not ppApp Is Nothing Then
ppApp.Visible = True
ppApp.Activate
With ppPres.SlideShowSettings
.ShowType = 1 ' ppShowTypeSpeaker (Full Screen)
.ShowPresenterView = 0 ' msoFalse (Disable presenter view)
.Run
End With
End If
On Error GoTo ERR_HANDLER"""

new_slide = """' Automatic slideshow removed (will prompt user instead at the end)"""

if target_slide in content:
    content = content.replace(target_slide, new_slide)
else:
    print("Could not find target_slide")

# Also need to postpone `Set ppPres = Nothing` and `Set ppApp = Nothing` until AFTER the message box.
target_cleanup = """910     Set ppPres = Nothing
920     Set ppApp = Nothing"""
new_cleanup = """' cleanup postponed"""
if target_cleanup in content:
    content = content.replace(target_cleanup, new_cleanup)

# --- Task 3: Modify the success message to a vbYesNoCancel prompt ---
target_msg = """Dim askMsg As String
If bSaved Then
askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1493) & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & vbCrLf & reportsFolder
MsgBoxU askMsg, vbOKOnly + vbInformation
Else
' "hamatzget notzra behatzlacha. yesh lishmor yadanit."
askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "." & vbCrLf & _
ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1513) & ChrW(1502) & ChrW(1493) & ChrW(1512) & " " & ChrW(1497) & ChrW(1491) & ChrW(1504) & ChrW(1497) & ChrW(1514) & "."
MsgBoxU askMsg, vbOKOnly + vbInformation
End If"""

new_msg = """Dim askMsg As String
Dim savedText As String
If bSaved Then
    savedText = ChrW(1493) & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " (" & reportsFolder & ")"
Else
    savedText = ChrW(1488) & ChrW(1498) & " " & ChrW(1506) & ChrW(1491) & ChrW(1497) & ChrW(1497) & ChrW(1503) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509)
End If

askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & ", " & savedText & vbCrLf & vbCrLf & _
         "האם תרצה להציג את המצגת כעת?" & vbCrLf & vbCrLf & _
         "- לחץ 'כן' (Yes) לתצוגת שקופיות (מסך מלא)" & vbCrLf & _
         "- לחץ 'לא' (No) לפתיחה במצב עריכה (PowerPoint)" & vbCrLf & _
         "- לחץ 'ביטול' (Cancel) כדי לסיים ללא הצגה"

Dim resMsg As Long
resMsg = MsgBoxU(askMsg, vbYesNoCancel + vbQuestion, "סיום הכנת מצגת")

On Error Resume Next
If Not ppApp Is Nothing And Not ppPres Is Nothing Then
    If resMsg = vbYes Then
        ppApp.Visible = True
        ppApp.Activate
        With ppPres.SlideShowSettings
            .ShowType = 1
            .ShowPresenterView = 0
            .Run
        End With
    ElseIf resMsg = vbNo Then
        ppApp.Visible = True
        ppApp.Activate
    Else
        ' Cancel
        If Not bSaved Then ppPres.Save
        ppPres.Close
        ppApp.Quit
    End If
End If
On Error GoTo ERR_HANDLER

910     Set ppPres = Nothing
920     Set ppApp = Nothing
"""

if target_msg in content:
    content = content.replace(target_msg, new_msg)
else:
    print("Could not find target_msg")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_120"', 'Attribute VB_Name = "Goren_Claude_V2_121"')
content = content.replace('VERSION: V2.120', 'VERSION: V2.121')
content = content.replace('APP_VERSION As String = "2.120"', 'APP_VERSION As String = "2.121"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.121 created.")
