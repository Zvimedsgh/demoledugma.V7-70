import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.023.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.024.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_023", "Goren_Claude_V2_024")
content = content.replace("V2.023", "V2.024")
content = content.replace('APP_VERSION As String = "2.023"', 'APP_VERSION As String = "2.024"')

# Force Years in BuildPresentation
old_pres = """20      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

        Dim yearVal As String
        Dim refYear As String"""
        
new_pres = """20      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

        Dim isDemoMode As Boolean
        Dim demoParam As String
        isDemoMode = False
        On Error Resume Next
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        On Error GoTo ERR_HANDLER

        Dim yearVal As String
        Dim refYear As String"""

content = content.replace(old_pres, new_pres)

old_pres2 = """40      yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
50      refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))"""

new_pres2 = """40      yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
50      refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
        
        If isDemoMode Then
            yearVal = "2025"
            refYear = "2024"
        End If"""

content = content.replace(old_pres2, new_pres2)

# Fix Title Position to EXACTLY E1
# Let's align to Left = wsMain.Range("E1").Left, Top = wsMain.Range("E1").Top. Width = 600.
old_title2 = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("E1").Left - 300, wsMain.Range("E1").Top, 600, 40)'''
new_title2 = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("G1").Left - 300, wsMain.Range("E1").Top, 600, 40)'''
# wait, if RTL, columns are decreasing leftwards. Let's just do Left = wsMain.Range("E1").Left
new_title3 = '''Set shpTitle = wsMain.Shapes.AddTextbox(1, wsMain.Range("E1").Left, wsMain.Range("E1").Top, 600, 40)'''

if old_title2 in content:
    content = content.replace(old_title2, new_title3)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.024 successfully")
