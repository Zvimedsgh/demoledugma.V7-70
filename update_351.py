import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.50.bas', 'r', encoding='utf-8') as f:
    text = f.read()

# Update version strings
text = text.replace('V3.50', 'V3.51')
text = text.replace('VERSION: 3.50', 'VERSION: 3.51')
text = text.replace('APP_VERSION As String = "3.50"', 'APP_VERSION As String = "3.51"')

# Update BuildPresentation to create presentation invisibly
build_pres_old = '''        If ppApp Is Nothing Then
            Set ppApp = CreateObject("PowerPoint.Application")
            ppWeOwnApp = True
        End If
615     ppApp.Visible = True
620     Set ppPres = ppApp.Presentations.Add
        On Error Resume Next
        ppApp.WindowState = 2 ' Minimized for speed and to keep Excel in focus
        AppActivate Application.Caption ' Give focus back to Excel
        On Error GoTo ERR_HANDLER'''

build_pres_new = '''        If ppApp Is Nothing Then
            Set ppApp = CreateObject("PowerPoint.Application")
            ppWeOwnApp = True
        End If
        ' Run presentation creation ENTIRELY IN THE BACKGROUND
620     Set ppPres = ppApp.Presentations.Add(0) ' 0 = msoFalse (invisible)
        On Error GoTo ERR_HANDLER'''

text = text.replace(build_pres_old, build_pres_new)

# Update end of BuildPresentation to show the window
end_sub_old = '''On Error Resume Next
If Not ppPres Is Nothing Then
    ppApp.Visible = True
    ppApp.WindowState = 3 ' Maximized
    ppPres.Slides(ppPres.Slides.Count).Select
    AppActivate ppApp.Caption
    
End If'''

end_sub_new = '''On Error Resume Next
If Not ppPres Is Nothing Then
    ppApp.Visible = True
    ppPres.NewWindow ' Create window for invisible presentation
    ppApp.WindowState = 3 ' Maximized
    ppPres.Slides(ppPres.Slides.Count).Select
    AppActivate ppApp.Caption
End If'''

text = text.replace(end_sub_old, end_sub_new)

# Update SafeExportChart to check for file existence
safe_export_old = '''Private Sub SafeExportChart(ByVal xlCht As Object, ByVal imgPath As String)
    Dim exportSuccess As Boolean
    DoEvents
    exportSuccess = xlCht.Export(imgPath)
    If Not exportSuccess Then
        DoEvents
        Application.Wait Now + TimeValue("00:00:01")
        exportSuccess = xlCht.Export(imgPath)
        If Not exportSuccess Then
            Err.Raise vbObjectError + 1, "SafeExportChart", "The specified file wasn't found (Chart.Export failed to save " & imgPath & ")"
        End If
    End If
End Sub'''

safe_export_new = '''Private Sub SafeExportChart(ByVal xlCht As Object, ByVal imgPath As String)
    Dim exportSuccess As Boolean
    Dim retryCount As Integer
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(imgPath) Then Kill imgPath
    DoEvents
    
    ' First attempt
    exportSuccess = xlCht.Export(imgPath)
    
    ' Wait loop: Ensure file is fully written and unlocked (Race Condition fix)
    Dim fileReady As Boolean
    fileReady = False
    retryCount = 0
    
    Do While Not fileReady And retryCount < 10 ' 10 attempts (up to 10 seconds)
        If fso.FileExists(imgPath) Then
            ' Check if file is readable
            On Error Resume Next
            Dim fileNum As Integer
            fileNum = FreeFile
            Open imgPath For Input Lock Read As #fileNum
            If Err.Number = 0 Then
                fileReady = True
            End If
            Close #fileNum
            On Error GoTo 0
        End If
        
        If Not fileReady Then
            Application.Wait Now + TimeValue("00:00:01")
            ' Try exporting again just in case the first one completely failed silently
            If Not exportSuccess Then exportSuccess = xlCht.Export(imgPath)
            retryCount = retryCount + 1
        End If
    Loop
    
    If Not fileReady Then
        Err.Raise vbObjectError + 1, "SafeExportChart", "The specified file wasn't found (Chart.Export failed to save " & imgPath & ")"
    End If
End Sub'''

text = text.replace(safe_export_old, safe_export_new)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Updated successfully')
