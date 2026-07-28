import re
with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.50.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('Attribute VB_Name = "Goren_Claude_V3_49"', 'Attribute VB_Name = "Goren_Claude_V3_50"')
text = text.replace("' VERSION: V3.49", "' VERSION: V3.50")
text = text.replace('Private Const APP_VERSION As String = "3.49"', 'Private Const APP_VERSION As String = "3.50"')

helper = '''\' ============================================================================
\' HELPER: Robust Chart Export with Retry
\' ============================================================================
Private Sub SafeExportChart(ByVal xlCht As Object, ByVal imgPath As String)
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
End Sub

\' ============================================================================'''

text = text.replace("' ============================================================================\n' HELPER: Export Total Summary chart to image file", helper + "\n' HELPER: Export Total Summary chart to image file")

text = re.sub(r'xlCht\.Export (img[a-zA-Z0-9_]+)', r'SafeExportChart xlCht, \1', text)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.50.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated successfully")
