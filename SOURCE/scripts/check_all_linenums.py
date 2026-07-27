import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
current_func = ""
line_nums = {}

for i, line in enumerate(lines):
    line_strip = line.strip()
    if line_strip.startswith("Private Sub") or line_strip.startswith("Public Sub") or line_strip.startswith("Function") or line_strip.startswith("Private Function") or line_strip.startswith("Public Function"):
        in_func = True
        current_func = line_strip
        line_nums = {}
    elif in_func and ("End Sub" in line_strip or "End Function" in line_strip):
        in_func = False
    elif in_func:
        parts = line_strip.split(maxsplit=1)
        if parts and parts[0].isdigit():
            ln = parts[0]
            if ln in line_nums:
                print(f"DUPLICATE LINE NUMBER IN {current_func}: {ln} at lines {line_nums[ln]} and {i+1}")
            else:
                line_nums[ln] = i+1

