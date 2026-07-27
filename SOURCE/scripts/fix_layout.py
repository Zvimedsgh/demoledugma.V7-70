import re
import sys

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.1.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 1. We need to swap F and G for Parameters, and J and K for Currency in the NEW block.
    # The new block has:
    # wsMain.Range("G2").Value = ChrW(1508) ... ' פרמטר
    # wsMain.Range("F2").Value = ChrW(1506) ... ' ערך
    # We want F=פרמטר, G=ערך.
    content = content.replace('wsMain.Range("G2").Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512)', 'wsMain.Range("F2").Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512)')
    content = content.replace('wsMain.Range("F2").Value = ChrW(1506) & ChrW(1512) & ChrW(1498)', 'wsMain.Range("G2").Value = ChrW(1506) & ChrW(1512) & ChrW(1498)')
    
    # Swap G3:G10 labels to F3:F10
    content = content.replace('wsMain.Range("G3").Value = ChrW(1513)', 'wsMain.Range("F3").Value = ChrW(1513)')
    content = content.replace('wsMain.Range("G4").Value = ChrW(1513)', 'wsMain.Range("F4").Value = ChrW(1513)')
    content = content.replace('wsMain.Range("G5").Value = ChrW(1505)', 'wsMain.Range("F5").Value = ChrW(1505)')
    content = content.replace('wsMain.Range("G6").Value = ChrW(1506)', 'wsMain.Range("F6").Value = ChrW(1506)')
    content = content.replace('wsMain.Range("G7").Value = ChrW(1505)', 'wsMain.Range("F7").Value = ChrW(1505)')
    content = content.replace('wsMain.Range("G8").Value = ChrW(1495)', 'wsMain.Range("F8").Value = ChrW(1495)')
    content = content.replace('wsMain.Range("G9").Value = ChrW(1506)', 'wsMain.Range("F9").Value = ChrW(1506)')
    content = content.replace('wsMain.Range("G10").Value = ChrW(1513)', 'wsMain.Range("F10").Value = ChrW(1513)')
    
    # Update styling ranges
    content = content.replace('wsMain.Range("G3:G10").Font.Color = RGB(0, 70, 140)', 'wsMain.Range("F3:F10").Font.Color = RGB(0, 70, 140)')
    content = content.replace('wsMain.Range("G3:G10").Font.Bold = True', 'wsMain.Range("F3:F10").Font.Bold = True')
    content = content.replace('wsMain.Range("G3:G10").HorizontalAlignment = xlRight', 'wsMain.Range("F3:F10").HorizontalAlignment = xlRight')
    
    # Swap Named Ranges F3:F10 to G3:G10
    content = content.replace('"rngBaseYear", wsMain.Range("F3")', '"rngBaseYear", wsMain.Range("G3")')
    content = content.replace('"rngCurrentYear", wsMain.Range("F4")', '"rngCurrentYear", wsMain.Range("G4")')
    content = content.replace('"rngPeriodType", wsMain.Range("F5")', '"rngPeriodType", wsMain.Range("G5")')
    content = content.replace('"rngPeriodValue", wsMain.Range("F6")', '"rngPeriodValue", wsMain.Range("G6")')
    content = content.replace('"rngDateType", wsMain.Range("F7")', '"rngDateType", wsMain.Range("G7")')
    content = content.replace('"rngFilterType", wsMain.Range("F8")', '"rngFilterType", wsMain.Range("G8")')
    content = content.replace('"rngFilterValue", wsMain.Range("F9")', '"rngFilterValue", wsMain.Range("G9")')
    content = content.replace('"rngClientName", wsMain.Range("F10")', '"rngClientName", wsMain.Range("G10")')
    
    # Swap default sets
    content = content.replace('IsEmpty(wsMain.Range("F5").Value) Then wsMain.Range("F5").Value', 'IsEmpty(wsMain.Range("G5").Value) Then wsMain.Range("G5").Value')
    content = content.replace('IsEmpty(wsMain.Range("F8").Value) Then wsMain.Range("F8").Value', 'IsEmpty(wsMain.Range("G8").Value) Then wsMain.Range("G8").Value')
    content = content.replace('IsEmpty(wsMain.Range("F10").Value) Then wsMain.Range("F10").Value', 'IsEmpty(wsMain.Range("G10").Value) Then wsMain.Range("G10").Value')
    
    # Swap Validation ranges
    content = content.replace('wsMain.Range("F5").Validation', 'wsMain.Range("G5").Validation')
    content = content.replace('wsMain.Range("F6").Validation', 'wsMain.Range("G6").Validation')
    content = content.replace('wsMain.Range("F7").Validation', 'wsMain.Range("G7").Validation')
    content = content.replace('wsMain.Range("F8").Validation', 'wsMain.Range("G8").Validation')
    content = content.replace('wsMain.Range("F9").Validation', 'wsMain.Range("G9").Validation')

    # Currency K->J, J->K
    content = content.replace('wsMain.Range("K2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)', 'wsMain.Range("J2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)')
    content = content.replace('wsMain.Range("J2").Value = ChrW(1513) & ChrW(1506) & ChrW(1512)', 'wsMain.Range("K2").Value = ChrW(1513) & ChrW(1506) & ChrW(1512)')
    content = content.replace('wsMain.Range("K3").Value = "$" & " "', 'wsMain.Range("J3").Value = "$" & " "')
    content = content.replace('wsMain.Range("K4").Value = ChrW(8364)', 'wsMain.Range("J4").Value = ChrW(8364)')
    content = content.replace('wsMain.Range("K3:K4").Font.Bold', 'wsMain.Range("J3:J4").Font.Bold')
    content = content.replace('wsMain.Range("K3:K4").HorizontalAlignment', 'wsMain.Range("J3:J4").HorizontalAlignment')
    
    content = content.replace('"rngDOLAR", wsMain.Range("J3")', '"rngDOLAR", wsMain.Range("K3")')
    
    # Rates were placed in J3/J4, move them to K3/K4
    content = content.replace('wsMain.Range("J3").Value = curRate', 'wsMain.Range("K3").Value = curRate')
    content = content.replace('IsEmpty(wsMain.Range("J3").Value)', 'IsEmpty(wsMain.Range("K3").Value)')
    content = content.replace('wsMain.Range("J3").Value = 3.6', 'wsMain.Range("K3").Value = 3.6')
    content = content.replace('wsMain.Range("J3").NumberFormat = "0.0000"', 'wsMain.Range("K3").NumberFormat = "0.0000"')
    
    content = content.replace('wsMain.Range("J4").Value = eurRate', 'wsMain.Range("K4").Value = eurRate')
    content = content.replace('IsEmpty(wsMain.Range("J4").Value)', 'IsEmpty(wsMain.Range("K4").Value)')
    content = content.replace('wsMain.Range("J4").Value = 3.9', 'wsMain.Range("K4").Value = 3.9')
    content = content.replace('wsMain.Range("J4").NumberFormat = "0.0000"', 'wsMain.Range("K4").NumberFormat = "0.0000"')
    content = content.replace('wsMain.Range("J3:J4").HorizontalAlignment = xlCenter', 'wsMain.Range("K3:K4").HorizontalAlignment = xlCenter')
    
    # Add green background
    content = content.replace('wsMain.Rows("11:13").RowHeight = 35', 'wsMain.Rows("11:13").RowHeight = 35\n    wsMain.Range("A1:U24").Interior.Color = RGB(220, 240, 220)')
    
    # Fix Helper buttons placement (F is right, G is left)
    # Search should be F, All should be G (or whatever matches screenshot).
    # In screenshot, "חפש" is on the right, "הכל" is on the left.
    # So "חפש" = F11, "הכל" = G11.
    content = content.replace('wsMain.Range("G11").Left + 5', 'wsMain.Range("F11").Left + 5')
    content = content.replace('wsMain.Range("G11").Top + 2', 'wsMain.Range("F11").Top + 2')
    content = content.replace('wsMain.Range("F11").Left + 5', 'wsMain.Range("G11").Left + 5')
    content = content.replace('wsMain.Range("F11").Top + 2', 'wsMain.Range("G11").Top + 2')
    
    # Reset button spans F and G
    # In RTL, F is right, G is left. So Left edge of G is the left edge of the whole block.
    # So Range("G12").Left is correct. Width = G12.Width + F12.Width.
    # This is already correct in the python script.
    
    # Now, REMOVE the obsolete blocks at the bottom of SetupMainSheet.
    # We want to remove from "    ' ---- Set default values if empty ----"
    # Down to "    ' ---- Set RTL and font size 14 for ALL sheets in workbook ----"
    
    start_str = "    ' ---- Set default values if empty ----\n"
    end_str = "    ' ---- Set RTL and font size 14 for ALL sheets in workbook ----\n"
    
    idx_start = content.rfind(start_str) # rfind to get the LAST occurrence (the old one)
    idx_end = content.find(end_str, idx_start)
    
    if idx_start != -1 and idx_end != -1:
        # Before deleting, wait, we must NOT delete the button deletion block!
        # Wait, the obsolete block has "btnSearchClient" delete, etc. We don't care, we already deleted them.
        content = content[:idx_start] + content[idx_end:]
        
    # Also remove the bottom duplicate blocks for borders and background:
    start_str2 = "    ' ---- Exchange rate info message at J5:K7 ----\n"
    end_str2 = "    ' ---- Clean old versions formatting ----\n"
    
    idx_start2 = content.rfind(start_str2)
    idx_end2 = content.find(end_str2, idx_start2)
    
    if idx_start2 != -1 and idx_end2 != -1:
        content = content[:idx_start2] + content[idx_end2:]
        
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
        
    print("Fixed layout script completed.")

if __name__ == "__main__":
    main()
