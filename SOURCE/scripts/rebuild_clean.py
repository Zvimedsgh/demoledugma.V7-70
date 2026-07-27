import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip_next = 0

for i, line in enumerate(lines):
    if skip_next > 0:
        skip_next -= 1
        continue
        
    # 1. Protect DATA_ sheets
    if "If sName = H_SET_PARAMS() Then IsProtectedSheet = True: Exit Function" in line:
        new_lines.append(line)
        new_lines.append("    If UCase$(Left$(sName, 5)) = \"DATA_\" Then IsProtectedSheet = True: Exit Function\n")
        continue

    # 2. Fix Demo Mode fallback - BuildReview
    if "If isDemoMode Then" in line and "BuildReview" in "".join(lines[max(0, i-50):i]):
        if "Set wsSrc = ThisWorkbook.Worksheets(\"DATA_\" & yearVal)" in lines[i+2]:
            new_lines.append('        If isDemoMode Then\n')
            new_lines.append('            On Error Resume Next\n')
            new_lines.append('            Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)\n')
            new_lines.append('            On Error GoTo ERR_HANDLER\n')
            new_lines.append('            If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"\n')
            new_lines.append('            Set wbSrc = ThisWorkbook\n')
            new_lines.append('        Else\n')
            new_lines.append('            srcPath = FindSourceFile(yearVal)\n')
            new_lines.append('540         If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal\n')
            new_lines.append('            Application.DisplayAlerts = True\n')
            new_lines.append('550         Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)\n')
            new_lines.append('            Application.DisplayAlerts = False\n')
            new_lines.append('560         Set wsSrc = OpenDataSheet(wbSrc)\n')
            new_lines.append('        End If\n')
            
            for j in range(i, i+30):
                if "Set wsSrc = OpenDataSheet(wbSrc)" in lines[j]:
                    if "End If" in lines[j+1]:
                        skip_next = (j - i) + 1
                        break
            continue

    # 3. Fix Demo Mode fallback - ApplyCorrectionsAndBuildReports (yearVal)
    if "If isDemoMode Then" in line and "ApplyCorrectionsAndBuildReports" in "".join(lines[max(0, i-150):i]):
        if "Set wsSrc = ThisWorkbook.Worksheets(\"DATA_\" & yearVal)" in lines[i+2]:
            new_lines.append('    If isDemoMode Then\n')
            new_lines.append('        On Error Resume Next\n')
            new_lines.append('        Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)\n')
            new_lines.append('        On Error GoTo ERR_HANDLER\n')
            new_lines.append('        If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"\n')
            new_lines.append('        Set wbSrc = ThisWorkbook\n')
            new_lines.append('    Else\n')
            new_lines.append('1050        srcPath = FindSourceFile(yearVal)\n')
            new_lines.append('1060        If srcPath = "" Then Err.Raise vbObjectError + 2002, "ApplyCorrections", "SOURCE NOT FOUND: " & yearVal\n')
            new_lines.append('1070        debugStep = "OPEN_SRC"\n')
            new_lines.append('1080        Set wbSrc = SafeOpenWorkbook(srcPath)\n')
            new_lines.append('        If wbSrc Is Nothing Then Err.Raise 424, "OPEN_SRC", "Failed to open source file. Excel might be blocking it."\n')
            new_lines.append('1090        Set wsSrc = OpenDataSheet(wbSrc)\n')
            new_lines.append('    End If\n')
            
            for j in range(i, i+30):
                if "Set wsSrc = OpenDataSheet(wbSrc)" in lines[j]:
                    if "End If" in lines[j+1]:
                        skip_next = (j - i) + 1
                        break
            continue

    # 4. Fix Demo Mode fallback - ApplyCorrectionsAndBuildReports (refYear)
    if "If isDemoMode Then" in line and "ApplyCorrectionsAndBuildReports" in "".join(lines[max(0, i-150):i]):
        if "Set wsRef = ThisWorkbook.Worksheets(\"DATA_\" & refYear)" in lines[i+2]:
            new_lines.append('    If isDemoMode Then\n')
            new_lines.append('        On Error Resume Next\n')
            new_lines.append('        Set wsRef = ThisWorkbook.Worksheets("DATA_" & refYear)\n')
            new_lines.append('        On Error GoTo ERR_HANDLER\n')
            new_lines.append('        If wsRef Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & refYear & " not found!"\n')
            new_lines.append('        Set wbRef = ThisWorkbook\n')
            new_lines.append('    Else\n')
            new_lines.append('1100        refPath = FindSourceFile(refYear)\n')
            new_lines.append('1110        If refPath = "" Then Err.Raise vbObjectError + 2003, "ApplyCorrections", "REF SOURCE NOT FOUND: " & refYear\n')
            new_lines.append('1120        debugStep = "OPEN_REF"\n')
            new_lines.append('1130        Set wbRef = SafeOpenWorkbook(refPath)\n')
            new_lines.append('        If wbRef Is Nothing Then Err.Raise 424, "OPEN_REF", "Failed to open ref file. Excel might be blocking it."\n')
            new_lines.append('1140        Set wsRef = OpenDataSheet(wbRef)\n')
            new_lines.append('    End If\n')
            
            for j in range(i, i+30):
                if "Set wsRef = OpenDataSheet(wbRef)" in lines[j]:
                    if "End If" in lines[j+1]:
                        skip_next = (j - i) + 1
                        break
            continue

    # 5. Fix chartItems = 0
    if "chartItems = nItems" in line:
        new_lines.append(line)
        new_lines.append("    If chartItems = 0 Then\n")
        new_lines.append("        Application.ScreenUpdating = True\n")
        new_lines.append("        Exit Sub\n")
        new_lines.append("    End If\n")
        continue

    # 6. Fix tmpWs.ChartObjects.Delete error throwing
    if "tmpWs.ChartObjects.Delete" in line:
        new_lines.append("        On Error Resume Next\n")
        new_lines.append(line)
        new_lines.append("        On Error GoTo ERR_HANDLER\n")
        continue

    new_lines.append(line)

# Now add line numbers to ExportCompCharts and ExportAgentCommissions
final_lines = []
in_func = False
line_num = 10
for i, line in enumerate(new_lines):
    if "Private Sub ExportCompCharts" in line or "Private Sub ExportAgentCommissions" in line:
        in_func = True
        line_num = 10
        final_lines.append(line)
        continue
    if in_func and "End Sub" in line:
        in_func = False
        final_lines.append(line)
        continue
        
    if in_func:
        stripped = line.lstrip()
        m = re.match(r'^(\d+)\s+(.*)', stripped)
        if m:
            stripped = m.group(2)
            
        if stripped != "" and not stripped.startswith("'") and not stripped.startswith("On Error") and not stripped.endswith(":") and "Dim " not in stripped and "ReDim " not in stripped:
            final_lines.append(f"{line_num} " + line.lstrip(' 0123456789'))
            line_num += 10
        else:
            final_lines.append(line)
    else:
        final_lines.append(line)


# Bump version
for i, line in enumerate(final_lines):
    if "Attribute VB_Name =" in line:
        final_lines[i] = line.replace("V2_050", "V2_056")
    if "Private Const APP_VERSION As String =" in line:
        final_lines[i] = line.replace("2.050", "2.056")
    if "VERSION: V2.050" in line:
        final_lines[i] = line.replace("2.050", "2.056")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_clean.bas', 'w', encoding='utf-8') as f:
    f.writelines(final_lines)

print("Created V2.056_clean")
