import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.028.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.029.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_028", "Goren_Claude_V2_029")
content = content.replace("V2.028", "V2.029")
content = content.replace('APP_VERSION As String = "2.028"', 'APP_VERSION As String = "2.029"')

# Force yearVal and refYear in BuildReview if DemoMode
old_br = """530     Dim isDemoMode As Boolean
        isDemoMode = False
        Dim demoParam As String
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        
        If isDemoMode Then
            On Error Resume Next
            Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)"""

new_br = """530     Dim isDemoMode As Boolean
        isDemoMode = False
        Dim demoParam As String
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        
        If isDemoMode Then
            yearVal = "2025"
            On Error Resume Next
            Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)"""

content = content.replace(old_br, new_br)

# Also force it if they run Button 1 for the refYear (sometimes they might run it twice? BuildReview doesn't use refYear)

# Wait, why did Button 2 use 2026 if I already forced it in V2.026/V2.027?
# Let's check ApplyCorrectionsAndBuildReports in V2.028
old_b2 = """        If isDemoMode Then
            yearVal = "2025"
            refYear = "2024"
        End If"""

new_b2 = """        If isDemoMode Then
            yearVal = "2025"
            refYear = "2024"
        End If"""
        
# It is already there! But wait, is there another place in ApplyCorrections where yearVal is read?
# Let's check!

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.029 successfully")
