Set xl = CreateObject("Excel.Application")
xl.Visible = False
Set wb = xl.Workbooks.Add
Set ws = wb.Worksheets(1)

' Unlock A1
ws.Range("A1").Locked = False

' Add validation
ws.Range("A1").Validation.Add 3, 1, 1, "1,2,3"

' Protect sheet
ws.Protect "password"

' Try to change validation
On Error Resume Next
ws.Range("A1").Validation.Delete
If Err.Number <> 0 Then
    WScript.Echo "Delete failed: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Delete succeeded!"
End If

ws.Range("A1").Validation.Add 3, 1, 1, "4,5,6"
If Err.Number <> 0 Then
    WScript.Echo "Add failed: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Add succeeded!"
End If

wb.Close False
xl.Quit
