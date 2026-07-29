import re
from datetime import datetime

file_path = r"C:\ledugma\DEMO\modDemoReports_V9.28.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Add ApplyDemoRestrictions sub at the end of the file
demo_restrictions_sub = """
' ============================================================================
' MACRO: ApplyDemoRestrictions
' Instantly updates the UI to show 2024/2025, blocks typing, and blocks clicks
' ============================================================================
Public Sub ApplyDemoRestrictions()
    On Error GoTo ERR_HANDLER
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))
    
    ' Set the values
    wsMain.Range("G3").Value = "2024"
    wsMain.Range("G4").Value = "2025"
    
    ' Block typing via Data Validation
    With wsMain.Range("G3:G4")
        .Validation.Delete
        .Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
        .Validation.ShowError = True
        .Validation.ErrorMessage = ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1500) & ChrW(1489) & ChrW(1491)
    End With
    
    ' Block clicking via transparent shape
    On Error Resume Next
    wsMain.Shapes("shpProYears").Delete
    On Error GoTo ERR_HANDLER
    
    Dim rngYearsBlock As Range
    Set rngYearsBlock = wsMain.Range("G3:G4")
    Dim shpYears As Shape
    Set shpYears = wsMain.Shapes.AddShape(msoShapeRectangle, rngYearsBlock.Left, rngYearsBlock.Top, rngYearsBlock.Width, rngYearsBlock.Height)
    shpYears.Name = "shpProYears"
    shpYears.Fill.Visible = msoFalse
    shpYears.Line.Visible = msoFalse
    shpYears.OnAction = "ProVersionOnly"
    
    MsgBox "UI updated to Demo restrictions (2024/2025).", vbInformation
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
"""

if "Public Sub ApplyDemoRestrictions" not in content:
    content += "\n" + demo_restrictions_sub

# Update SetupMainSheet to call it instead of doing it inline
old_inline = """        ' ---- Override years to 2024/2025 and block UI with transparent shape ----
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        On Error Resume Next
        wsMain.Shapes("shpProYears").Delete
        On Error GoTo ERR_HANDLER
        Dim rngYearsBlock As Range
        Set rngYearsBlock = wsMain.Range("G3:G4")
        Dim shpYears As Shape
        Set shpYears = wsMain.Shapes.AddShape(msoShapeRectangle, rngYearsBlock.Left, rngYearsBlock.Top, rngYearsBlock.Width, rngYearsBlock.Height)
        shpYears.Name = "shpProYears"
        shpYears.Fill.Visible = msoFalse
        shpYears.Line.Visible = msoFalse
        shpYears.OnAction = "ProVersionOnly" """

new_inline = """        ' ---- Override years to 2024/2025 and block UI with transparent shape ----
        ' Instead of doing it inline, we call the dedicated sub
        ' We don't want the MsgBox to popup during SetupMainSheet, so we do it silently
        wsMain.Range("G3").Value = "2024"
        wsMain.Range("G4").Value = "2025"
        With wsMain.Range("G3:G4")
            .Validation.Delete
            .Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
            .Validation.ShowError = True
            .Validation.ErrorMessage = ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1500) & ChrW(1489) & ChrW(1491)
        End With
        On Error Resume Next
        wsMain.Shapes("shpProYears").Delete
        On Error GoTo ERR_HANDLER
        Dim rngYearsBlock As Range
        Set rngYearsBlock = wsMain.Range("G3:G4")
        Dim shpYears As Shape
        Set shpYears = wsMain.Shapes.AddShape(msoShapeRectangle, rngYearsBlock.Left, rngYearsBlock.Top, rngYearsBlock.Width, rngYearsBlock.Height)
        shpYears.Name = "shpProYears"
        shpYears.Fill.Visible = msoFalse
        shpYears.Line.Visible = msoFalse
        shpYears.OnAction = "ProVersionOnly" """

content = content.replace(old_inline, new_inline)

# Update version header
now = datetime.now()
date_str = now.strftime("%Y-%m-%d %H:%M:%S")

content = re.sub(r"VERSION: 9\.27 \(.*?\)", f"VERSION: 9.28 ({date_str})", content)
content = content.replace("VERSION 9.27", "VERSION 9.28")
content = content.replace("CHANGES IN 9.27:", "CHANGES IN 9.28:\n'   - Added ApplyDemoRestrictions to instantly apply UI blocks\n'   - Added Data Validation to block manual typing via keyboard in G3:G4\n' CHANGES IN 9.27:")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Done patch5")
