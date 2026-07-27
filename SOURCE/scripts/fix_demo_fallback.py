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
        
    # In BuildReview:
    if "If isDemoMode Then" in line and "Set wsSrc =" in lines[i+2] and "DATA_" in lines[i+2]:
        # We replace the block:
        # If isDemoMode Then
        # On Error Resume Next
        # Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)
        # On Error GoTo ERR_HANDLER
        # If Not wsSrc Is Nothing Then Set wbSrc = ThisWorkbook
        # End If
        
        # If wbSrc Is Nothing Then
        new_lines.append('If isDemoMode Then\n')
        new_lines.append('    On Error Resume Next\n')
        new_lines.append('    Set wsSrc = ThisWorkbook.Worksheets("DATA_" & yearVal)\n')
        new_lines.append('    On Error GoTo ERR_HANDLER\n')
        new_lines.append('    If wsSrc Is Nothing Then Err.Raise vbObjectError + 3000, "DemoMode", "Demo sheet DATA_" & yearVal & " not found!"\n')
        new_lines.append('    Set wbSrc = ThisWorkbook\n')
        new_lines.append('Else\n')
        new_lines.append('    srcPath = FindSourceFile(yearVal)\n')
        new_lines.append('    If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal\n')
        new_lines.append('    Application.DisplayAlerts = True\n')
        new_lines.append('    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)\n')
        new_lines.append('    Application.DisplayAlerts = False\n')
        new_lines.append('End If\n')
        
        # We need to skip until the end of the original `If wbSrc Is Nothing` block
        j = i + 1
        while j < len(lines):
            if "Application.DisplayAlerts = False" in lines[j] and "wbSrc" in lines[j-1]:
                skip_next = (j - i) + 1  # Skip the End If as well, wait, is there an End If?
                if "End If" in lines[j+1]:
                    skip_next = (j - i) + 2
                break
            j += 1
        continue

    # Same logic for ApplyCorrectionsAndBuildReports (yearVal and refYear)
    
    # Bump version
    if "Attribute VB_Name =" in line:
        new_lines.append(line.replace("V2_050", "V2_051"))
        continue
    if "Private Const APP_VERSION As String =" in line:
        new_lines.append(line.replace("2.050", "2.051"))
        continue
    if "VERSION: V2.050" in line:
        new_lines.append(line.replace("2.050", "2.051"))
        continue

    new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.051.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.051")
