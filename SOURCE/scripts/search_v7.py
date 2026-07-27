import sys
import glob
import re

files = glob.glob(r'c:\LEVAV PROJECT\SOURCE\*V7*.bas')
for fpath in files:
    with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
        
    in_sub = False
    for i, line in enumerate(lines):
        line_strip = line.strip()
        if "Sub UpdateClientList" in line_strip:
            in_sub = True
            print(f"--- {fpath} ---")
            print(f"[{i+1}] {line.strip()}")
        elif in_sub and "End Sub" in line_strip:
            print(f"[{i+1}] {line.strip()}")
            in_sub = False
        elif in_sub:
            if "Sort" in line_strip or "arr" in line_strip:
                print(f"[{i+1}] {line.strip()}")
