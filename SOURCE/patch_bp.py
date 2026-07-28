import re

def patch():
    with open(r'C:\LEVAV PROJECT\SOURCE\old_bp_clean.txt', 'r', encoding='utf-8') as f:
        lines = [line.rstrip('\n') for line in f.readlines()]
    
    # We will replace lines 117 to 323 with our new logic.
    # First, let's find the exact indices.
    # We look for: '        ' ================================================================'
    #              '        ' PHASE 1: Create all chart images in Excel (NO PowerPoint yet)'
    
    start_idx = -1
    end_idx = -1
    for i, line in enumerate(lines):
        if 'PHASE 1: Create all chart images in Excel' in line:
            start_idx = i - 1 # include the '====' above it
            break
            
    for i in range(start_idx, len(lines)):
        if 'PHASE 3:' in line:
            pass # Keep looking
        if 'Add page numbers to all slides' in lines[i]:
            end_idx = i - 1 # include the empty line above it
            break

    if start_idx == -1 or end_idx == -1:
        print(f"Error: could not find indices. start={start_idx}, end={end_idx}")
        return

    replacement = '''
        ' ================================================================
        ' PHASE 1 & 2: Open PowerPoint and Build Slides via Clipboard
        ' ================================================================
        Dim ppApp As Object
        Dim ppPres As Object
        Dim ppSlide As Object
        Dim ppWeOwnApp As Boolean
        ppWeOwnApp = False
        On Error Resume Next
        Set ppApp = GetObject(, "PowerPoint.Application")
        On Error GoTo ERR_HANDLER
        If ppApp Is Nothing Then
            Set ppApp = CreateObject("PowerPoint.Application")
            ppWeOwnApp = True
        End If
        ppApp.Visible = True
        Set ppPres = ppApp.Presentations.Add
        On Error Resume Next
        ppApp.WindowState = 2 ' Minimized for speed and to keep Excel in focus
        AppActivate Application.Caption ' Give focus back to Excel
        On Error GoTo ERR_HANDLER

        ' Set LANDSCAPE slide size (13.33" x 7.5")
        ppPres.PageSetup.SlideWidth = 960
        ppPres.PageSetup.SlideHeight = 540

        Dim slideIdx As Long
        Dim slideW As Single
        Dim slideH As Single
        slideIdx = 0
        slideW = 960
        slideH = 540

        ' Slide title names (Hebrew)
        Dim titleNames(1 To 6) As String
        titleNames(1) = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
        titleNames(2) = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
        titleNames(3) = H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
        titleNames(4) = H_TELLER() & ChrW(1497) & ChrW(1493) & ChrW(1514)
        titleNames(5) = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)

        ' SLIDE 1: Title
        slideIdx = slideIdx + 1
        Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
        BuildTitleSlide ppSlide, yearVal, refYear, periodDesc, slideW, slideH, paramsSubtitle

        ' SLIDE 2: Total Summary chart
        BuildTotalSlide ppPres, yearVal, refYear, slideW, paramsSubtitle
        slideIdx = ppPres.Slides.Count ' Sync slide count

        ' Build list of sheets to process
        Dim sheetList(1 To 6) As String
        Dim sheetCount As Long
        sheetCount = 0

        If SheetExists(SHEET_MONTHS()) Then
            sheetCount = sheetCount + 1
            sheetList(sheetCount) = SHEET_MONTHS()
        End If
        If SheetExists(SHEET_COMPANIES()) Then
            sheetCount = sheetCount + 1
            sheetList(sheetCount) = SHEET_COMPANIES()
        End If
        If SheetExists(SHEET_MAINBRANCH()) Then
            sheetCount = sheetCount + 1
            sheetList(sheetCount) = SHEET_MAINBRANCH()
        End If
        If SheetExists(SHEET_TELLERS()) Then
            sheetCount = sheetCount + 1
            sheetList(sheetCount) = SHEET_TELLERS()
        End If
        If SheetExists(SHEET_AGENTS()) Then
            sheetCount = sheetCount + 1
            sheetList(sheetCount) = SHEET_AGENTS()
        End If

        ' For each comparison sheet: build 4 chart slides + 1 table slide
        Dim si As Long
        For si = 1 To sheetCount
            ' This generates 4 slides directly inside PowerPoint!
            BuildCompSlides ppPres, sheetList(si), titleNames(si), yearVal, refYear, slideW, paramsSubtitle
            
            slideIdx = ppPres.Slides.Count
            
            ' Slide: Data table (always)
            slideIdx = slideIdx + 1
            Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
            BuildTableSlide ppSlide, sheetList(si), titleNames(si), yearVal, refYear, slideW, slideH, paramsSubtitle
        Next si

        ' ================================================================
        ' PHASE 3: "Agents without Levav" variant slides (2 chart slides)
        ' ================================================================
        If SheetExists(SHEET_AGENTS()) Then
            Dim levavName As String
            levavName = GetActiveAgencyName()
            
            ' Generate the NO-LEVAV charts:
            BuildCompSlides ppPres, SHEET_AGENTS(), titleNames(5), yearVal, refYear, slideW, paramsSubtitle, levavName, False, False
            
            slideIdx = ppPres.Slides.Count
        End If
'''
    # We also need to fix the ERR_HANDLER to not try to kill image files!
    # And we need to remove the image killing code before restoring screen updating.
    
    new_lines = lines[:start_idx] + replacement.split('\n') + lines[end_idx+1:]
    
    # Remove kill logic
    # Find "Cleanup temp images" up to "Re-hide sheets"
    kill_start = -1
    kill_end = -1
    for i, line in enumerate(new_lines):
        if "' Cleanup temp images" in line:
            kill_start = i
        if "' Re-hide sheets that were hidden before" in line:
            kill_end = i - 1
            break
            
    if kill_start != -1 and kill_end != -1:
        new_lines = new_lines[:kill_start] + new_lines[kill_end+1:]
        
    # Also in ERR_HANDLER
    err_kill_start = -1
    err_kill_end = -1
    for i, line in enumerate(new_lines):
        if "If imgTotal <>" in line:
            err_kill_start = i - 2
        if "End If" in line and "User-friendly error message" in new_lines[i+2]:
            err_kill_end = i + 1
            break
            
    if err_kill_start != -1 and err_kill_end != -1:
        new_lines = new_lines[:err_kill_start] + new_lines[err_kill_end+1:]

    with open(r'C:\LEVAV PROJECT\SOURCE\new_bp_fixed.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(new_lines))
        
    print("Patched BuildPresentation saved to new_bp_fixed.txt")

if __name__ == '__main__':
    patch()
