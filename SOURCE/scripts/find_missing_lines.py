import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.037_20260702_1430.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_func = True
    if in_func and "End Sub" in line:
        break
    if in_func:
        # Check if line does NOT start with a line number
        stripped = line.strip()
        if stripped and not stripped.startswith("'") and not stripped.startswith("Dim ") and not stripped.startswith("Public Sub"):
            # Check if it starts with a digit
            if not stripped[0].isdigit():
                print(f"[{i}] {stripped}")
