Set app = CreateObject("Excel.Application")
app.Visible = False
Set wb = app.Workbooks.Add()
Set ws = wb.Sheets(1)
ws.Range("G6:H6").Merge
Set target = ws.Range("G6:H6")
Set G6 = ws.Range("G6")
Set intersect = app.Intersect(target, G6)
If intersect Is Nothing Then
    WScript.Echo "Nothing"
Else
    WScript.Echo "Not Nothing, Address: " & intersect.Address
End If
wb.Close False
app.Quit
