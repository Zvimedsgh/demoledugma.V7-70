import sys
import shutil

old_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.025.bas'
new_filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.026.bas'

shutil.copy(old_filepath, new_filepath)

with open(new_filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace("Goren_Claude_V2_025", "Goren_Claude_V2_026")
content = content.replace("V2.025", "V2.026")
content = content.replace('APP_VERSION As String = "2.025"', 'APP_VERSION As String = "2.026"')

# Fix duplicates in ApplyCorrectionsAndBuildReports
old_dup1 = """1040    debugStep = "CHECK_SRC"
        Dim isDemoMode As Boolean
        isDemoMode = FORCE_DEMO_MODE
        Dim demoParam As String
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True"""

new_dup1 = """1040    debugStep = "CHECK_SRC"
        isDemoMode = FORCE_DEMO_MODE
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True"""

content = content.replace(old_dup1, new_dup1)

old_dup2 = """        Dim isRefDemoMode As Boolean
        isRefDemoMode = FORCE_DEMO_MODE
        Dim refDemoParam As String
        refDemoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If refDemoParam = ChrW(1499) & ChrW(1503) Or refDemoParam = "YES" Then isRefDemoMode = True
        
        If isRefDemoMode Then"""

new_dup2 = """        ' Reusing isDemoMode for Ref source
        isDemoMode = FORCE_DEMO_MODE
        demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
        If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True
        
        If isDemoMode Then"""

content = content.replace(old_dup2, new_dup2)

with open(new_filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.026 successfully")
