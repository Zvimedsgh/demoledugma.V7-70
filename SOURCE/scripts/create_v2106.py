import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.105.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Let's add the UI update logic directly to the end of CheckUserPermissions!
# We don't need to call A00_SetupMainSheet. We just need to:
# 1. Update the Title text
# 2. Update the Title color based on Demo Mode
# 3. ApplyDemoLockOnOpen (which we know works safely)
# 4. Jump to F3

def update_permissions_end(match):
    return """
    ' --- Update Home Page UI on Open ---
    On Error Resume Next
    Dim isDemoModeUI As Boolean
    Dim demoParamUI As String
    isDemoModeUI = FORCE_DEMO_MODE
    demoParamUI = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If Not FORCE_DEMO_MODE Then
        If demoParamUI = ChrW(1499) & ChrW(1503) Or demoParamUI = "YES" Then isDemoModeUI = True
        If demoParamUI = ChrW(1500) & ChrW(1488) Or demoParamUI = "NO" Then isDemoModeUI = False
    End If
    
    Dim wsMainUI As Worksheet
    Set wsMainUI = ThisWorkbook.Worksheets(homeSheetName)
    
    wsMainUI.Unprotect "Z961814r"
    wsMainUI.Range("A1").Value = GetActiveAgencyName()
    
    If isDemoModeUI Then
        wsMainUI.Range("A1").Font.Color = RGB(200, 0, 0) ' Red for Demo
    Else
        wsMainUI.Range("A1").Font.Color = RGB(0, 0, 0) ' Black for Real
    End If
    
    ApplyDemoLockOnOpen
    Application.Goto wsMainUI.Range("F3")
    wsMainUI.Protect Password:="Z961814r", UserInterfaceOnly:=True
    On Error GoTo 0
    ' -----------------------------------

    ' Navigate to home page
    ThisWorkbook.Worksheets(homeSheetName).Activate
"""

content = re.sub(r'\n\s*\' Navigate to home page\s*\n\s*ThisWorkbook\.Worksheets\(homeSheetName\)\.Activate', update_permissions_end, content)

# Also fix A00_SetupMainSheet to NOT always make the text red!
def fix_setupmain_color(match):
    return """
    If isDemoMode Then
        wsMain.Range("A1").Font.Color = RGB(200, 0, 0) ' Red for Demo
    Else
        wsMain.Range("A1").Font.Color = RGB(0, 0, 0) ' Black for Real
    End If
    """
content = re.sub(r'\n\s*3460 wsMain\.Range\("A1"\)\.Font\.Color = RGB\(200, 0, 0\)', fix_setupmain_color, content)

# Remove Auto_Open
content = re.sub(r'Public Sub Auto_Open\(\).*?End Sub\s*\n', '', content, flags=re.DOTALL)

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_105"', 'Attribute VB_Name = "Goren_Claude_V2_106"')
content = content.replace('VERSION: V2.105', 'VERSION: V2.106')
content = content.replace('APP_VERSION As String = "2.105"', 'APP_VERSION As String = "2.106"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.106.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.106 created.")
