Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open("c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.69.xlsm")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("c:\LEVAV PROJECT\SOURCE\wsMain_V1.69.cls", True)
objFile.Write objWorkbook.VBProject.VBComponents("wsMain").CodeModule.Lines(1, objWorkbook.VBProject.VBComponents("wsMain").CodeModule.CountOfLines)
objFile.Close
objWorkbook.Close False
objExcel.Quit
