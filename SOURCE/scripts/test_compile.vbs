Set xl = CreateObject("Excel.Application")
xl.Visible = False
Set wb = xl.Workbooks.Add
On Error Resume Next
wb.VBProject.VBComponents.Import "c:\LEVAV PROJECT\SOURCE\Levav_Manus _V7.51.bas"
If Err.Number <> 0 Then
    WScript.Echo "Import Error: " & Err.Description
    xl.Quit
    WScript.Quit 1
End If
On Error GoTo 0
WScript.Echo "Imported successfully."
xl.Quit
