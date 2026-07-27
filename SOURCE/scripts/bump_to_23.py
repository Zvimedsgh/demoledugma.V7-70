import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.022.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.023.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update version
content = content.replace("Goren_Claude_V2_022", "Goren_Claude_V2_023")
content = content.replace("V2.022", "V2.023")
content = content.replace('APP_VERSION As String = "2.022"', 'APP_VERSION As String = "2.023"')

# 2. Fix Title Position to E1
old_title = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, 100, 5, 800, 40)'''
new_title = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("I1").Left, wsMain.Range("E1").Top, wsMain.Range("B1").Left - wsMain.Range("I1").Left, 40)'''
# wait, if RTL, B1.Left > I1.Left. Let's just use 600 width.
new_title2 = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("E1").Left - 300, wsMain.Range("E1").Top, 600, 40)'''

if old_title in content:
    content = content.replace(old_title, new_title2)

# 3. Force Years in ApplyCorrectionsAndBuildReports
old_years = """600     yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
610     refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
620     If yearVal = "" Or refYear = "" Then Err.Raise vbObjectError + 2001, "ApplyCorrections", "B2 OR B3 IS EMPTY"
"""
new_years = """600     yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
610     refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
        
        Dim isDemoMode As Boolean
        Dim demoParam As String
        isDemoMode = False
        On Error Resume Next
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        On Error GoTo ERR_HANDLER
        
        If isDemoMode Then
            yearVal = "2025"
            refYear = "2024"
        End If

620     If yearVal = "" Or refYear = "" Then Err.Raise vbObjectError + 2001, "ApplyCorrections", "B2 OR B3 IS EMPTY"
"""
content = content.replace(old_years, new_years)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.023 successfully")
