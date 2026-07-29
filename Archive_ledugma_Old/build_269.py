import os

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.268.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('APP_VERSION As String = "2.268"', 'APP_VERSION As String = "2.269"')
content = content.replace('VERSION: V2.268', 'VERSION: V2.269')
content = content.replace('Error in V2.268!', 'Error in V2.269!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_268"', 'Attribute VB_Name = "Goren_Claude_V2_269"')

# Changelog
changelog = """' CHANGES IN 2.269:
'   - UI: Repositioned all stray elements for Laptop mode (Manual button, Toggle Sheets, Exit button) to fit perfectly on small screens.
'   - UI: Cleared leftover V2.257 text from previous layouts.
"""
content = content.replace("' CHANGES IN 2.268:", changelog + "' CHANGES IN 2.268:")


# Fix ToggleDisplayMode to handle ALL the laptop layout elements
old_toggle_laptop = """        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("G13")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("G14")
        
        Dim shpManual As Shape
        Set shpManual = wsMain.Shapes("shpMatachManual")
        If Not shpManual Is Nothing Then
            shpManual.Left = wsMain.Range("A13").Left + 60
            shpManual.Top = wsMain.Range("A13").Top
        End If"""

new_toggle_laptop = """        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("G13")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("G14")
        
        ' 1. Matach Manual Button
        Dim shpManual As Shape
        Set shpManual = wsMain.Shapes("shpMatachManual")
        If Not shpManual Is Nothing Then
            shpManual.Left = wsMain.Range("F15").Left + 20
            shpManual.Top = wsMain.Range("F15").Top + 5
        End If
        
        ' 2. Operation Manual Button (shpManualMsg)
        Dim shpOpManual As Shape
        Set shpOpManual = wsMain.Shapes("shpManualMsg")
        If Not shpOpManual Is Nothing Then
            shpOpManual.Left = wsMain.Range("D12").Left
            shpOpManual.Top = wsMain.Range("D12").Top
        End If
        
        ' 3. Toggle Sheets Button
        Dim shpToggle As Shape
        Set shpToggle = wsMain.Shapes("shpToggleSheets")
        If Not shpToggle Is Nothing Then
            shpToggle.Left = wsMain.Range("A12").Left + 60
            shpToggle.Top = wsMain.Range("A12").Top
        End If
        
        ' 4. Exit Button
        Dim shpExitBtn As Shape
        Set shpExitBtn = wsMain.Shapes("btnNavExit")
        If Not shpExitBtn Is Nothing Then
            shpExitBtn.Left = wsMain.Range("A15").Left + 60
            shpExitBtn.Top = wsMain.Range("A15").Top
        End If
        
        ' 5. Clear old V2.257 texts from bottom
        wsMain.Range("A20:Z30").ClearContents"""

content = content.replace(old_toggle_laptop, new_toggle_laptop)


old_toggle_desktop = """        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("K4")
        
        Dim shpManual2 As Shape
        Set shpManual2 = wsMain.Shapes("shpMatachManual")
        If Not shpManual2 Is Nothing Then
            shpManual2.Left = wsMain.Range("L1").Left + (wsMain.Range("L1").Width - 120) / 2
            shpManual2.Top = wsMain.Range("L1").Top + wsMain.Range("L1").Height - 80
        End If"""

new_toggle_desktop = """        ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
        ThisWorkbook.Names.Add "rngEURO", wsMain.Range("K4")
        
        ' 1. Matach Manual Button
        Dim shpManual2 As Shape
        Set shpManual2 = wsMain.Shapes("shpMatachManual")
        If Not shpManual2 Is Nothing Then
            shpManual2.Left = wsMain.Range("L1").Left + (wsMain.Range("L1").Width - 120) / 2
            shpManual2.Top = wsMain.Range("L1").Top + wsMain.Range("L1").Height - 80
        End If
        
        ' 2. Operation Manual Button (shpManualMsg)
        Dim shpOpManual2 As Shape
        Set shpOpManual2 = wsMain.Shapes("shpManualMsg")
        If Not shpOpManual2 Is Nothing Then
            shpOpManual2.Left = wsMain.Range("L2").Left + (wsMain.Range("L2").Width - 120) / 2
            shpOpManual2.Top = wsMain.Range("L2").Top + 20
        End If
        
        ' 3. Toggle Sheets Button
        Dim shpToggle2 As Shape
        Set shpToggle2 = wsMain.Shapes("shpToggleSheets")
        If Not shpToggle2 Is Nothing Then
            shpToggle2.Left = wsMain.Range("A14").Left + 60
            shpToggle2.Top = wsMain.Range("A14").Top
        End If
        
        ' 4. Exit Button
        Dim shpExitBtn2 As Shape
        Set shpExitBtn2 = wsMain.Shapes("btnNavExit")
        If Not shpExitBtn2 Is Nothing Then
            shpExitBtn2.Left = wsMain.Range("A19").Left + 60
            shpExitBtn2.Top = wsMain.Range("A19").Top
        End If
        
        ' 5. Clear old V2.257 texts from bottom
        wsMain.Range("A20:Z30").ClearContents"""

content = content.replace(old_toggle_desktop, new_toggle_desktop)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.269.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 269')
