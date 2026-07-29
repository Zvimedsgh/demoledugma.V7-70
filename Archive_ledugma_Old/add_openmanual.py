with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

macro = """
' ============================================================================
' OPEN MANUAL
' ============================================================================
Sub OpenManual()
    On Error Resume Next
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Manual.pdf"
    
    If Dir(pdfPath) <> "" Then
        ActiveWorkbook.FollowHyperlink pdfPath
    Else
        MsgBoxU ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & ChrW(1492) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " (Manual.pdf) " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1489) & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & ".", vbCritical
    End If
    On Error GoTo 0
End Sub
"""

if "Sub OpenManual()" not in text:
    text += macro
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
        f.write(text)
    print("Macro OpenManual added.")
else:
    print("Macro OpenManual already exists.")
