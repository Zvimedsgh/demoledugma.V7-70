import sys
import os

folder = r'c:\LEVAV PROJECT\SOURCE'
for root, dirs, files in os.walk(folder):
    for file in files:
        if file.endswith('.bas') or file.endswith('.vba'):
            filepath = os.path.join(root, file)
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                if "הוראות_תפעול" in content and ("Add" in content or "AddShape" in content):
                    # check if we actually CREATE the sheet
                    if "Sheets.Add" in content or "Worksheets.Add" in content:
                        pass
