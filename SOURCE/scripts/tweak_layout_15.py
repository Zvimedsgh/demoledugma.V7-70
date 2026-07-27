import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.34.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Update SetupSectionNamedRanges to handle offsetRow
    old_setup_names = """    ' --- Create section Named Ranges ---
    ThisWorkbook.Names.Add "rngSection_FieldMap", wsMgmt.Range("A1")
    ThisWorkbook.Names.Add "rngSection_BranchName", wsMgmt.Range("A57")
    ThisWorkbook.Names.Add "rngSection_Params", wsMgmt.Range("A173")
    ThisWorkbook.Names.Add "rngSection_ReasonCode", wsMgmt.Range("A181")
    ThisWorkbook.Names.Add "rngSection_PeriodLists", wsMgmt.Range("A196")
    ThisWorkbook.Names.Add "rngSection_Messages", wsMgmt.Range("A221")
    ThisWorkbook.Names.Add "rngSection_Permissions", wsMgmt.Range("A240")
    ThisWorkbook.Names.Add "rngSection_Clients", wsMgmt.Range("A249")
    
    ' --- Create folder path Named Ranges (Params section) ---
    ThisWorkbook.Names.Add "rngFILES_FOLDER", wsMgmt.Range("B176")
    ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsMgmt.Range("B177")"""

    new_setup_names = """    Dim offsetRow As Long
    offsetRow = 0
    If wsMgmt.Range("A1").Value = "MENU_AREA" Then offsetRow = 4

    ' --- Create section Named Ranges ---
    ThisWorkbook.Names.Add "rngSection_FieldMap", wsMgmt.Range("A" & 1 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_BranchName", wsMgmt.Range("A" & 57 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_Params", wsMgmt.Range("A" & 173 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_ReasonCode", wsMgmt.Range("A" & 181 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_PeriodLists", wsMgmt.Range("A" & 196 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_Messages", wsMgmt.Range("A" & 221 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_Permissions", wsMgmt.Range("A" & 240 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_Clients", wsMgmt.Range("A" & 249 + offsetRow)
    
    ' --- Create folder path Named Ranges (Params section) ---
    ThisWorkbook.Names.Add "rngFILES_FOLDER", wsMgmt.Range("B" & 176 + offsetRow)
    ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsMgmt.Range("B" & 177 + offsetRow)"""

    content = content.replace(old_setup_names, new_setup_names)

    # 2. Update SetupSettingsMenu to arrange buttons in rows 1-2 and freeze panes
    old_setup_menu = """    ' Remove old menu buttons
  140 On Error Resume Next
  150 For Each s In wsMgmt.Shapes
  160     If Left$(s.Name, 6) = "btnNav" Then s.Delete
  170 Next s
  180 On Error GoTo 0
      
      ' Button dimensions - placed in column I
  190 btnLeft = wsMgmt.Range("I1").Left + 5
  200 btnTop = wsMgmt.Range("I2").Top
  210 btnW = 140
  220 btnH = 26
  230 btnGap = btnH + 4
      
      ' Title in I1
  240 wsMgmt.Range("I1").Value = ChrW(1504) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1496)  ' "?????"
  250 wsMgmt.Range("I1").Font.Size = 12
  260 wsMgmt.Range("I1").Font.Bold = True
  270 wsMgmt.Range("I1").Font.Color = RGB(0, 0, 102)"""

    new_setup_menu = """    ' Insert top rows for toolbar if not already present
      If wsMgmt.Range("A1").Value <> "MENU_AREA" Then
          wsMgmt.Rows("1:4").Insert Shift:=xlDown
          wsMgmt.Rows("1:4").ClearFormats
          wsMgmt.Rows("1:4").Interior.Color = RGB(240, 245, 250) ' light blue-gray
          wsMgmt.Range("A1").Value = "MENU_AREA"
          wsMgmt.Range("A1").Font.Color = RGB(240, 245, 250) ' hidden
          
          ' Set freeze panes
          wsMgmt.Activate
          ActiveWindow.FreezePanes = False
          ActiveWindow.SplitRow = 4
          ActiveWindow.SplitColumn = 0
          ActiveWindow.FreezePanes = True
      End If

    ' Remove old menu buttons
  140 On Error Resume Next
  150 For Each s In wsMgmt.Shapes
  160     If Left$(s.Name, 6) = "btnNav" Then s.Delete
  170 Next s
  180 On Error GoTo 0
      
      ' Button dimensions - placed horizontally
  210 btnW = 140
  220 btnH = 26
  230 btnGap = btnW + 10
      
      ' Title in J1 (moved from I1)
  240 wsMgmt.Range("J1").Value = ChrW(1504) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1496)  ' "?????"
  250 wsMgmt.Range("J1").Font.Size = 12
  260 wsMgmt.Range("J1").Font.Bold = True
  270 wsMgmt.Range("J1").Font.Color = RGB(0, 0, 102)"""

    content = content.replace(old_setup_menu, new_setup_menu)

    # Now replace the btnLeft and btnTop logic for each button
    # btnNavFieldMap
    content = content.replace("280 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)",
                              "    btnLeft = 10\n    btnTop = 5\n280 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)")
    content = content.replace("370 btnTop = btnTop + btnGap", "370 btnLeft = btnLeft + btnGap")

    # btnNavBranchName
    content = content.replace("470 btnTop = btnTop + btnGap", "470 btnLeft = btnLeft + btnGap")

    # btnNavParams
    content = content.replace("570 btnTop = btnTop + btnGap", "570 btnLeft = btnLeft + btnGap")

    # btnNavReasonCode (this is the 4th button, next should wrap to row 2)
    content = content.replace("670 btnTop = btnTop + btnGap", "670 btnLeft = 10\n    btnTop = 35")

    # btnNavPeriodType
    content = content.replace("770 btnTop = btnTop + btnGap", "770 btnLeft = btnLeft + btnGap")

    # btnNavMessages
    content = content.replace("870 btnTop = btnTop + btnGap", "870 btnLeft = btnLeft + btnGap")

    # btnNavPermissions
    content = content.replace("970 btnTop = btnTop + btnGap", "970 btnLeft = btnLeft + btnGap")

    # Write back
    with open("c:\\LEVAV PROJECT\\SOURCE\\tweak_layout_15.py.out", 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == "__main__":
    main()
