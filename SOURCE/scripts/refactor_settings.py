import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Fix SettingsGoToSheet
new_settings_goto = """Private Sub SettingsGoToSheet(shtName As String)
On Error Resume Next
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets(shtName)
If Not ws Is Nothing Then
If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
ws.Activate
Application.Goto ws.Range("A1")
End If
End Sub"""

content = re.sub(
    r'Private Sub SettingsGoToSheet\(shtName As String\).*?End Sub',
    new_settings_goto,
    content,
    flags=re.DOTALL
)

# 2. Fix SetupMainSheet Parameters logic
# Need to replace `wsMgmt` with `wsParams` in the parameter finding section
# The param section starts around `paramStartRow = 3` and goes to the backup check
def replace_params(match):
    text = match.group(0)
    text = text.replace('wsMgmt', 'wsParams')
    return f"""Dim wsParams As Worksheet
On Error Resume Next
Set wsParams = ThisWorkbook.Worksheets(H_SET_PARAMS())
On Error GoTo ERR_HANDLER
If Not wsParams Is Nothing Then
{text}
End If
"""

content = re.sub(
    r'(paramStartRow = 3.*?2120  wsMgmt\.Cells\(paramLastRow \+ 1, COL_PARAM_VALUE\)\.Value = "C:\\LEVAV PROJECT\\BACKUPS")',
    replace_params,
    content,
    flags=re.DOTALL
)

# 3. Fix DEMO_MODE read before the param setup block
content = re.sub(
    r'(isDemoMode = FORCE_DEMO_MODE\s*On Error Resume Next\s*)demoParam = UCase\$\(Trim\$\(GetStringParameter\(wsMgmt, "DEMO_MODE"\)\)\)',
    r'\1demoParam = UCase$(Trim$(GetStringParameter(ThisWorkbook.Worksheets(H_SET_PARAMS()), "DEMO_MODE")))',
    content
)

# 4. Fix BuildReview wsMgmt -> wsBranches
# In BuildReview, we have `Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())`
# Change it to `Set wsBranches = ThisWorkbook.Worksheets(H_SET_BRANCHES())` and update usage
build_review_pattern = r'(Public Sub BuildReview\(\).*?End Sub)'
def fix_build_review(match):
    text = match.group(1)
    # Replace wsMgmt definition
    text = text.replace('Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())', 'Set wsBranches = ThisWorkbook.Worksheets(H_SET_BRANCHES())')
    text = text.replace('Dim wsMgmt As Worksheet', 'Dim wsBranches As Worksheet')
    # Replace wsMgmt usage related to branches (around line 949)
    text = text.replace('wsMgmt.Cells(wsMgmt.Rows.Count', 'wsBranches.Cells(wsBranches.Rows.Count')
    text = text.replace('wsMgmt.Cells(brLastRow', 'wsBranches.Cells(brLastRow')
    text = text.replace('wsMgmt.Cells(brScan', 'wsBranches.Cells(brScan')
    text = text.replace('wsMgmt.Rows(brEndRow)', 'wsBranches.Rows(brEndRow)')
    return text

content = re.sub(build_review_pattern, fix_build_review, content, flags=re.DOTALL)

# 5. Fix ApplyCorrections wsMgmt -> wsBranches
apply_corr_pattern = r'(Public Sub ApplyCorrectionsAndBuildReports\(\).*?End Sub)'
def fix_apply_corr(match):
    text = match.group(1)
    text = text.replace('Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())', 'Set wsBranches = ThisWorkbook.Worksheets(H_SET_BRANCHES())')
    text = text.replace('Dim wsMgmt As Worksheet', 'Dim wsBranches As Worksheet')
    text = text.replace('LoadCorrectionsToDicts wsMgmt,', 'LoadCorrectionsToDicts wsBranches,')
    return text

content = re.sub(apply_corr_pattern, fix_apply_corr, content, flags=re.DOTALL)

# 6. Fix LoadCorrectionsToDicts signature and usage
load_corr_pattern = r'(Private Sub LoadCorrectionsToDicts\(ByVal wsMgmt As Worksheet.*?End Sub)'
def fix_load_corr(match):
    text = match.group(1)
    text = text.replace('ByVal wsMgmt As Worksheet', 'ByVal wsBranches As Worksheet')
    text = text.replace('wsMgmt.Cells(', 'wsBranches.Cells(')
    text = text.replace('wsMgmt.Rows.Count', 'wsBranches.Rows.Count')
    return text

content = re.sub(load_corr_pattern, fix_load_corr, content, flags=re.DOTALL)


# Update the version
content = content.replace('Private Const APP_VERSION As String = "2.035"', 'Private Const APP_VERSION As String = "2.037"')
content = content.replace('Private Const APP_VERSION As String = "2.036"', 'Private Const APP_VERSION As String = "2.037"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_fixed.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Created Goren_Claude_V2.037_fixed.bas")
