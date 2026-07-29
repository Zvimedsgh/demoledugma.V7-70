Set xl = CreateObject("Excel.Application")
xl.Visible = False
Set wb = xl.Workbooks.Add
Set ws = wb.Worksheets(1)

' Case 1: Sheet is unprotected
On Error Resume Next
ws.Unprotect "wrongpwd"
If Err.Number <> 0 Then
    WScript.Echo "Unprotected sheet Unprotect error: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Unprotected sheet Unprotect success!"
End If

' Case 2: Sheet is protected with NO password
ws.Protect ""
ws.Unprotect "wrongpwd"
If Err.Number <> 0 Then
    WScript.Echo "No-pwd sheet Unprotect error: " & Err.Description
    Err.Clear
Else
    WScript.Echo "No-pwd sheet Unprotect success!"
End If

wb.Close False
xl.Quit
