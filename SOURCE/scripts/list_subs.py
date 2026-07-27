import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.082.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    line_strip = line.strip()
    if line_strip.startswith("Public Sub ") or line_strip.startswith("Private Sub ") or line_strip.startswith("Function "):
        print(f"[{i+1}] {line_strip}")

