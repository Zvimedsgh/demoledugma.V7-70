import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1 & 4: HideWorkSheets and DisplayAlerts

# In ApplyCorrectionsAndBuildReports, add HideWorkSheets before success message
target_apply = """2795    MsgBoxU ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501), vbInformation"""
replacement_apply = """Call HideWorkSheets
2795    MsgBoxU ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501), vbInformation"""
if target_apply in content:
    content = content.replace(target_apply, replacement_apply)

# In BuildPresentation, add DisplayAlerts = False before Delete and HideWorkSheets before End Sub / MsgBox
# Let's find ws.Delete in BuildPresentation
old_del = """            If SheetExists(ws.Name) Then ws.Delete
        End If
    Next ws
    
    ' Build title slide"""
new_del = """            If SheetExists(ws.Name) Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next ws
    
    ' Build title slide"""
if old_del in content:
    content = content.replace(old_del, new_del)

target_pres = """MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1493) & ChrW(1492) & ChrW(1511) & ChrW(1489) & ChrW(1510) & ChrW(1497) & ChrW(1501) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1493) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & _"""
replacement_pres = """Call HideWorkSheets
    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1493) & ChrW(1492) & ChrW(1511) & ChrW(1489) & ChrW(1510) & ChrW(1497) & ChrW(1501) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1493) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!" & vbCrLf & _"""
if target_pres in content:
    content = content.replace(target_pres, replacement_pres)

# Fix 3: Remove refYear hardcode in ApplyCorrections
bad_refyear = """If isDemoMode Then
refYear = "2024"
End If"""
if bad_refyear in content:
    content = content.replace(bad_refyear, "")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119_temp.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Temp created.")
