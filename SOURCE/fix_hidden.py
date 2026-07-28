with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas', 'r', encoding='utf-8') as f:
    content = f.read()

missing_logic = '''
        ' Unhide all sheets before accessing data (in case some are VeryHidden)
        Dim wsTmp3 As Worksheet
        Dim hiddenSheets3() As String
        Dim hiddenCount3 As Long
        hiddenCount3 = 0
        ReDim hiddenSheets3(1 To ThisWorkbook.Worksheets.Count)
        For Each wsTmp3 In ThisWorkbook.Worksheets
            If wsTmp3.Visible <> xlSheetVisible Then
                hiddenCount3 = hiddenCount3 + 1
                hiddenSheets3(hiddenCount3) = wsTmp3.Name
                wsTmp3.Visible = xlSheetVisible
            End If
        Next wsTmp3

        ' ================================================================
        ' PHASE 1 & 2: Open PowerPoint and Build Slides via Clipboard
'''

content = content.replace("' ================================================================\n        ' PHASE 1 & 2: Open PowerPoint and Build Slides via Clipboard", missing_logic)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.51.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Fixed missing hiddenCount3 logic')
