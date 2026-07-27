import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.020.bas'
with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Replace the first 10 lines with the correct header
new_header = """Attribute VB_Name = "Goren_Claude_V2_020"
' ============================================================================
' MODULE: modLevav
' PURPOSE: Complete system - BuildReview + ApplyCorrectionsAndBuildReports
' VERSION: V2.020
' - UI: Agency Name Support, Demo Mode parameters, UI locking.
' DATE: 2026-07-02 11:00
' ============================================================================
"""

# Find where the actual code or previous header ends
# It usually starts with ' CHANGES IN 1.51:
start_idx = 0
for i, line in enumerate(lines):
    if "' CHANGES IN 1.51:" in line:
        start_idx = i
        break

if start_idx == 0:
    start_idx = 8 # fallback

lines = [new_header] + lines[start_idx:]

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Fixed headers in V2.020")
