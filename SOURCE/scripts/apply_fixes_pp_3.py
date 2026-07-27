import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_msg = "Dim askMsg As String"
start_idx2 = content.find(target_msg)
if start_idx2 != -1:
    end_str2 = "End If\n\nApplication.EnableEvents = True"
    end_idx2 = content.find(end_str2, start_idx2)
    if end_idx2 != -1:
        block2 = content[start_idx2:end_idx2+6]
        new_msg = """Dim askMsg As String
Dim savedText As String
If bSaved Then
    savedText = ChrW(1493) & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " (" & reportsFolder & ")"
Else
    savedText = ChrW(1488) & ChrW(1498) & " " & ChrW(1506) & ChrW(1491) & ChrW(1497) & ChrW(1497) & ChrW(1503) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509)
End If

askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & ", " & savedText & vbCrLf & vbCrLf & _
         ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1514) & ChrW(1512) & ChrW(1510) & ChrW(1492) & " " & ChrW(1500) & ChrW(1492) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1499) & ChrW(1506) & ChrW(1514) & "?" & vbCrLf & vbCrLf & _
         "- " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " 'כן' (Yes) " & ChrW(1500) & ChrW(1514) & ChrW(1510) & ChrW(1493) & ChrW(1490) & ChrW(1514) & " " & ChrW(1513) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " (" & ChrW(1502) & ChrW(1505) & ChrW(1498) & " " & ChrW(1502) & ChrW(1500) & ChrW(1488) & ")" & vbCrLf & _
         "- " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " 'לא' (No) " & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1492) & " " & ChrW(1489) & ChrW(1502) & ChrW(1510) & ChrW(1489) & " " & ChrW(1506) & ChrW(1512) & ChrW(1497) & ChrW(1499) & ChrW(1492) & " (PowerPoint)" & vbCrLf & _
         "- " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " 'ביטול' (Cancel) " & ChrW(1499) & ChrW(1491) & ChrW(1497) & " " & ChrW(1500) & ChrW(1505) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1492) & ChrW(1510) & ChrW(1490) & ChrW(1492)

Dim resMsg As Long
resMsg = MsgBoxU(askMsg, vbYesNoCancel + vbQuestion, ChrW(1505) & ChrW(1497) & ChrW(1493) & ChrW(1501) & " " & ChrW(1492) & ChrW(1499) & ChrW(1504) & ChrW(1514) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514))

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
        content = content.replace(block2, new_msg)
        print("Replaced target_msg")
    else:
        print("End str2 not found")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.121 msgbox replaced.")
