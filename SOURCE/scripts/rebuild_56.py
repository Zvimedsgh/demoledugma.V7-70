import sys

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
    if "If isDemoMode Then" in line and "BuildReview" in "".join(lines[i-50:i]):
        if "Set wsSrc = ThisWorkbook.Worksheets(\"DATA_\" & yearVal)" in lines[i+2]:
            new_lines.append('    If isDemoMode Then\n')
            new_lines.append('        On Error Resume Next\n')
            new_lines.append('        Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)\n')
            new_lines.append('        On Error GoTo ERR_HANDLER\n')
            new_lines.append('        If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"\n')
            new_lines.append('        Set wbSrc = ThisWorkbook\n')
            new_lines.append('    Else\n')
            new_lines.append('        srcPath = FindSourceFile(yearVal)\n')
            new_lines.append('        If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal\n')
            new_lines.append('        Application.DisplayAlerts = True\n')
            new_lines.append('        Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)\n')
            new_lines.append('        Application.DisplayAlerts = False\n')
            new_lines.append('    End If\n')
            
            # Skip the old block
            # Old block:
            # If isDemoMode Then
            # On Error Resume Next
            # Set wsSrc = ...
            # On Error GoTo ERR_HANDLER
            # If Not wsSrc Is Nothing Then Set wbSrc = ThisWorkbook
            # End If
            # 
            # If wbSrc Is Nothing Then
            # srcPath = FindSourceFile(yearVal)
            # If srcPath = "" Then Err.Raise ...
            # Application.DisplayAlerts = True
            # Set wbSrc = Workbooks.Open(...)
            # Application.DisplayAlerts = False
            # End If
            
            # We skip 14 lines. Let's verify by checking lines[i+14]
            skip_next = 13
            if "End If" not in lines[i+13]:
                # Dynamic skip:
                for j in range(i, i+30):
                    if "Application.DisplayAlerts = False" in lines[j]:
                        if "End If" in lines[j+1]:
                            skip_next = (j - i) + 1
                            break
            continue

    # 3. Fix Demo Mode fallback - ApplyCorrectionsAndBuildReports (yearVal)
    if "If isDemoMode Then" in line and "ApplyCorrectionsAndBuildReports" in "".join(lines[i-150:i]):
        if "Set wsSrc = ThisWorkbook.Worksheets(\"DATA_\" & yearVal)" in lines[i+2]:
            new_lines.append('    If isDemoMode Then\n')
            new_lines.append('        On Error Resume Next\n')
            new_lines.append('        Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)\n')
            new_lines.append('        On Error GoTo ERR_HANDLER\n')
            new_lines.append('        If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"\n')
            new_lines.append('        Set wbSrc = ThisWorkbook\n')
            new_lines.append('    Else\n')
            new_lines.append('        srcPath = FindSourceFile(yearVal)\n')
            new_lines.append('        If srcPath = "" Then Err.Raise vbObjectError + 2002, "ApplyCorrections", "SOURCE NOT FOUND: " & yearVal\n')
            new_lines.append('        debugStep = "OPEN_SRC"\n')
            new_lines.append('        Set wbSrc = SafeOpenWorkbook(srcPath)\n')
            new_lines.append('        If wbSrc Is Nothing Then Err.Raise 424, "OPEN_SRC", "Failed to open source file. Excel might be blocking it."\n')
            new_lines.append('        Set wsSrc = OpenDataSheet(wbSrc)\n')
            new_lines.append('    End If\n')
            
            for j in range(i, i+30):
                if "Set wsSrc = OpenDataSheet(wbSrc)" in lines[j]:
                    if "End If" in lines[j+1]:
                        skip_next = (j - i) + 1
                        break
            continue

    # 4. Fix Demo Mode fallback - ApplyCorrectionsAndBuildReports (refYear)
    if "If isDemoMode Then" in line and "ApplyCorrectionsAndBuildReports" in "".join(lines[i-150:i]):
        if "Set wsRef = ThisWorkbook.Worksheets(\"DATA_\" & refYear)" in lines[i+2]:
            new_lines.append('    If isDemoMode Then\n')
            new_lines.append('        On Error Resume Next\n')
            new_lines.append('        Set wsRef = ThisWorkbook.Worksheets("DATA_" & refYear)\n')
            new_lines.append('        On Error GoTo ERR_HANDLER\n')
            new_lines.append('        If wsRef Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & refYear & " not found!"\n')
            new_lines.append('        Set wbRef = ThisWorkbook\n')
            new_lines.append('    Else\n')
            new_lines.append('        refPath = FindSourceFile(refYear)\n')
            new_lines.append('        If refPath = "" Then Err.Raise vbObjectError + 2003, "ApplyCorrections", "REF SOURCE NOT FOUND: " & refYear\n')
            new_lines.append('        debugStep = "OPEN_REF"\n')
            new_lines.append('        Set wbRef = SafeOpenWorkbook(refPath)\n')
            new_lines.append('        If wbRef Is Nothing Then Err.Raise 424, "OPEN_REF", "Failed to open ref file. Excel might be blocking it."\n')
            new_lines.append('        Set wsRef = OpenDataSheet(wbRef)\n')
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

# Bump version
for i, line in enumerate(new_lines):
    if "Attribute VB_Name =" in line:
        new_lines[i] = line.replace("V2_050", "V2_056")
    if "Private Const APP_VERSION As String =" in line:
        new_lines[i] = line.replace("2.050", "2.056")
    if "VERSION: V2.050" in line:
        new_lines[i] = line.replace("2.050", "2.056")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.056")
