
On Error Resume Next
Set ppApp = CreateObject("PowerPoint.Application")
Set ppPres = ppApp.Presentations.Add(0)
Set slide = ppPres.Slides.Add(1, 12)
slide.Shapes.AddPicture "C:\Windows\Web\Wallpaper\Windows\img0.jpg", 0, 1, 10, 10, 100, 100
if Err.Number <> 0 Then
    WScript.Echo "Error AddPicture: " & Err.Description
End If

