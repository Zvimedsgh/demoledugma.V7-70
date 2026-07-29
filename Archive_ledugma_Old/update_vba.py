import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Add global variables at the top
cache_vars = '''
Private Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"
Private g_CachedUserAuth As Integer ' 0=uninit, 1=true, -1=false
'''
content = content.replace('Private Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"', cache_vars)

# 2. Add Helper Functions
helpers = '''
' ============================================================================
' HELPER FUNCTIONS FOR DEMO AND PERMISSIONS
' ============================================================================
Public Function IsUserAuthorized() As Boolean
    If g_CachedUserAuth = 1 Then IsUserAuthorized = True: Exit Function
    If g_CachedUserAuth = -1 Then IsUserAuthorized = False: Exit Function
    
    Dim winUser As String
    winUser = LCase$(Trim$(Environ("USERNAME")))
    If Len(winUser) = 0 Then
        g_CachedUserAuth = -1
        IsUserAuthorized = False
        Exit Function
    End If
    
    Dim wsMgmt As Worksheet
    On Error Resume Next
    Set wsMgmt = ThisWorkbook.Worksheets(H_SET_PERMISSIONS())
    If wsMgmt Is Nothing Then
        g_CachedUserAuth = -1
        IsUserAuthorized = False
        Exit Function
    End If
    
    Dim r As Long
    Dim cellVal As String
    Dim accessLevel As String
    r = 2
    Do While True
        cellVal = Trim$(CStr(wsMgmt.Cells(r, 1).Value2))
        If UCase$(cellVal) = "EOD" Or cellVal = "" Then Exit Do
        If LCase$(cellVal) = winUser Then
            accessLevel = Trim$(CStr(wsMgmt.Cells(r, 2).Value2))
            If UCase$(accessLevel) = "FULLACCESS" Or accessLevel = ChrW(1502) & ChrW(1500) & ChrW(1488) Then
                g_CachedUserAuth = 1
                IsUserAuthorized = True
                Exit Function
            End If
        End If
        r = r + 1
    Loop
    g_CachedUserAuth = -1
    IsUserAuthorized = False
End Function

Public Function IsSystemInDemoMode() As Boolean
    If FORCE_DEMO_MODE Then
        IsSystemInDemoMode = True
        Exit Function
    End If
    
    ' If user is NOT authorized, force demo mode!
    If Not IsUserAuthorized() Then
        IsSystemInDemoMode = True
        Exit Function
    End If
    
    ' If user IS authorized, check the parameter sheet
    Dim demoParamStr As String
    On Error Resume Next
    demoParamStr = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))
    If demoParamStr = ChrW(1499) & ChrW(1503) Or demoParamStr = "YES" Or demoParamStr = "" Then
        IsSystemInDemoMode = True
    Else
        IsSystemInDemoMode = False
    End If
End Function

'''
# Insert before CheckUserPermissions
content = content.replace("' ============================================================================\n' CHECK USER PERMISSIONS", helpers + "' ============================================================================\n' CHECK USER PERMISSIONS")

# 3. Replace GetActiveAgencyName
agency_func = '''Public Function GetActiveAgencyName() As String
    Application.Volatile
    If IsSystemInDemoMode() Then
        GetActiveAgencyName = Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), PARAM_DEMO_AGENCY_NAME))
    Else
        GetActiveAgencyName = Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), PARAM_AGENCY_NAME))
    End If
End Function'''
pattern_agency = re.compile(r'Public Function GetActiveAgencyName.*?End Function', re.DOTALL)
content = pattern_agency.sub(agency_func, content)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.13.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Done step 1')
