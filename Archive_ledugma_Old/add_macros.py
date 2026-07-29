with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

macros = """
' ============================================================================
' HIDE INSTRUCTIONS
' ============================================================================
Sub HideInstructions()
    On Error Resume Next
    Dim wsHome As Worksheet
    Set wsHome = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    ThisWorkbook.Unprotect "Z961814r"
    wsHome.Unprotect "Z961814r"
    
    wsHome.Range("AA1").Value = "YES"
    
    wsHome.Protect Password:="Z961814r", UserInterfaceOnly:=True
    ThisWorkbook.Protect "Z961814r"
    
    Call HideWorkSheets
    On Error GoTo 0
End Sub

' ============================================================================
' OPEN INSTRUCTIONS
' ============================================================================
Sub OpenInstructions()
    On Error Resume Next
    Dim wsHome As Worksheet
    Set wsHome = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    ThisWorkbook.Unprotect "Z961814r"
    wsHome.Unprotect "Z961814r"
    
    wsHome.Range("AA1").Value = "NO"
    
    wsHome.Protect Password:="Z961814r", UserInterfaceOnly:=True
    ThisWorkbook.Protect "Z961814r"
    
    Call HideWorkSheets
    ThisWorkbook.Worksheets(ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492)).Activate
    On Error GoTo 0
End Sub
"""

if "Sub HideInstructions()" not in text:
    text += macros
    with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
        f.write(text)
    print("Macros added.")
else:
    print("Macros already exist.")
