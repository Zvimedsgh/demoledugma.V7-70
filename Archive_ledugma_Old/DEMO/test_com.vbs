On Error Resume Next
Set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: " & Err.Description
Else
    WScript.Echo "Success! " & xl.ActiveWorkbook.Name
End If
