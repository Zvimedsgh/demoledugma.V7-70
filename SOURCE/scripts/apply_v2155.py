import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version numbers
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_153"', 'Attribute VB_Name = "Goren_Claude_V2_155"')
content = content.replace('VERSION: V2.153', 'VERSION: V2.155')
content = content.replace('APP_VERSION As String = "2.153"', 'APP_VERSION As String = "2.155"')

# Remove the buggy threshold filter from BuildBaseSheet
old_filter = "1440                If Abs(premVal) > threshold And Not dictCorrections.Exists(CStr(r)) Then GoTo NextSrcRow"
new_filter = "1440                ' REMOVED: threshold filter (Bug from 2.153 that deleted all data if parameter was missing)"

if old_filter in content:
    content = content.replace(old_filter, new_filter)
    print("Successfully removed threshold filter.")
else:
    print("ERROR: Could not find threshold filter!")

# Add Data Validation back to OpenClientSearch
# We need to find where wsSearch.Cells(2, 1) is formatted, around line 6168.
old_search_ui = """    ' Row 2: Yellow search cell
    wsSearch.Cells(2, 1).Interior.Color = RGB(255, 255, 200)
    wsSearch.Cells(2, 1).Font.Size = 14
    wsSearch.Cells(2, 1).Borders.LineStyle = xlContinuous
    wsSearch.Columns(1).ColumnWidth = 40"""

new_search_ui = """    ' Row 2: Yellow search cell with Dropdown (Data Validation) + Wildcard support
    wsSearch.Cells(2, 1).Interior.Color = RGB(255, 255, 200)
    wsSearch.Cells(2, 1).Font.Size = 14
    wsSearch.Cells(2, 1).Borders.LineStyle = xlContinuous
    wsSearch.Columns(1).ColumnWidth = 40
    
    ' Add Data Validation pointing to lst_clients, but allow typing any text (ShowError = False)
    On Error Resume Next
    With wsSearch.Cells(2, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_clients"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
    On Error GoTo 0"""

if old_search_ui in content:
    content = content.replace(old_search_ui, new_search_ui)
    print("Successfully added Data Validation back to OpenClientSearch.")
else:
    print("ERROR: Could not find old_search_ui block!")
    
with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.155 generated successfully.")
