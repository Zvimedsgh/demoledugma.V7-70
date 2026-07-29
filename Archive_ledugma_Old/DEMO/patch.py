import re

file_path = r"C:\ledugma\DEMO\modDemoReports_V7.95.bas"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Replace year assignment in BuildReview and ApplyCorrectionsAndBuildReports
# They look like:
# 600     yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
# 610     refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
content = re.sub(
    r'(yearVal = Trim\$\(CStr\(wsMain\.Range\("rngCurrentYear"\)\.Value2\)\))',
    r'yearVal = "2020" \' Forced for demo\n        \' \1',
    content
)
content = re.sub(
    r'(refYear = Trim\$\(CStr\(wsMain\.Range\("rngBaseYear"\)\.Value2\)\))',
    r'refYear = "2019" \' Forced for demo\n        \' \1',
    content
)

# In BuildPresentation, we have:
#        ' Force year to 2025/2024 for demo output
#        yearVal = "2025"
#        refYear = "2024"
content = content.replace('yearVal = "2025"', 'yearVal = "2020"')
content = content.replace('refYear = "2024"', 'refYear = "2019"')

# Add SetupMainSheet shape logic right before ' ---- Remove Page Break lines
setup_injection = """        ' ---- Override years to 2024/2025 and block UI with transparent shape ----
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
        shpYears.OnAction = "ProVersionOnly"

        ' ---- Remove Page Break lines"""
content = content.replace("        ' ---- Remove Page Break lines", setup_injection)

# Update version header
content = content.replace("VERSION: 7.94", "VERSION: 7.95")
content = content.replace("CHANGES IN 7.94:", "CHANGES IN 7.95:\n'   - Blocked UI years (2024/2025) and forced backend processing to 2019/2020\n' CHANGES IN 7.94:")

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Done")
