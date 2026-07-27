import sys
import datetime

filepath_old = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.035.bas'
filepath_new = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath_old, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace the specific header lines
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_035"', 'Attribute VB_Name = "Goren_Claude_V2_036"')
content = content.replace("' VERSION: V2.035", "' VERSION: V2.036")
content = content.replace("' DATE: 2026-07-02 11:00", "' DATE: 2026-07-02 12:57")
content = content.replace("' - UI: Agency Name Support, Demo Mode parameters, UI locking.", "' - UI: Agency Name Support, Demo Mode parameters, UI locking.\n' CHANGES IN 2.036:\n'   - FIX: Data Validation Native Formulas for Dropdowns, Demo Lock fixes.")

with open(filepath_new, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"Created {filepath_new}")
