import re
with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas','r',encoding='utf-8') as f:
    text = f.read()

# 1. Modify HideWorkSheets to protect at the end
# It currently has: ThisWorkbook.Worksheets(ctrlName).Activate \n End Sub
hide_new = """
' Activate home sheet
ThisWorkbook.Worksheets(ctrlName).Activate
' --- ADMIN MODE: LOCK SYSTEM ---
On Error Resume Next
ThisWorkbook.Worksheets(ctrlName).Protect Password:="Z961814r", UserInterfaceOnly:=True
ThisWorkbook.Protect Password:="Z961814r"
On Error GoTo 0
End Sub
"""
text = re.sub(r"' Activate home sheet\s*ThisWorkbook\.Worksheets\(ctrlName\)\.Activate\s*End Sub", hide_new.strip(), text, flags=re.DOTALL)

# 2. Modify ShowHiddenSheets to unprotect at the beginning
show_old = """Sub ShowHiddenSheets()
    ' Shows all hidden sheets except MATACH
    Dim ws As Worksheet
    On Error Resume Next"""
show_new = """Sub ShowHiddenSheets()
    ' Shows all hidden sheets except MATACH
    ' --- ADMIN MODE: UNLOCK SYSTEM ---
    On Error Resume Next
    ThisWorkbook.Unprotect Password:="Z961814r"
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Unprotect Password:="Z961814r"
    On Error GoTo 0
    
    Dim ws As Worksheet
    On Error Resume Next"""
text = text.replace(show_old, show_new)

# 3. Add SetupAdminProtection()
setup_sub = """
' ============================================================================
' ADMIN MODE SETUP (Run once to configure sheet protection)
' ============================================================================
Sub SetupAdminProtection()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    On Error Resume Next
    ThisWorkbook.Unprotect Password:="Z961814r"
    ws.Unprotect Password:="Z961814r"
    On Error GoTo 0
    
    ' Lock all cells by default
    ws.Cells.Locked = True
    
    ' Unlock specific input cells (Base Year, Current Year, Period Type, etc.)
    ' Assuming these are in B3:C4, E3:F6, E8 based on UI
    On Error Resume Next
    ws.Range("B3:C4").Locked = False
    ws.Range("E3:F6").Locked = False
    ws.Range("E8").Locked = False
    ws.Range("C8").Locked = False
    
    ' Also unlock any buttons (shapes are locked by default if they have macro, but let's ensure they work)
    ' Actually, Protect with UserInterfaceOnly allows macros to run, but clicking buttons works anyway 
    ' if the shapes aren't explicitly locked against clicking, which they aren't.
    
    MsgBoxU ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1504) & ChrW(1506) & ChrW(1497) & ChrW(1500) & ChrW(1492) & " " & ChrW(1492) & ChrW(1493) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
End Sub

"""
text = text + setup_sub

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas','w',encoding='utf-8') as f:
    f.write(text)

print("Modifications applied successfully.")
