import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056_20260702_1706.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub BuildPresentation" in line or "Private Sub BuildPresentation" in line:
        print(f"[{i+1}] {line.strip()}")
    if "Public Sub ExportCompCharts" in line or "Private Sub ExportCompCharts" in line:
        print(f"[{i+1}] {line.strip()}")

