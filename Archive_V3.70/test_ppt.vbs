
Set ppApp = CreateObject("PowerPoint.Application")
Set ppPres = ppApp.Presentations.Add(0)
ppPres.PageSetup.SlideWidth = 960
ppPres.PageSetup.SlideHeight = 540
Set slide = ppPres.Slides.Add(1, 12)
WScript.Sleep 2000
ppApp.Visible = True
ppPres.NewWindow

