import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.094.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "eurRate > 0" in line or "IsEmpty(wsMain.Range" in line or "wsMain.Range" in line and "K4" in line:
        pass
        
    if "K4" in line and "Value =" in line:
        # We will write it to a file so we can read it safely without encoding issues
        with open(r'c:\LEVAV PROJECT\SOURCE\k4_line.txt', 'a') as outf:
            outf.write(line)

