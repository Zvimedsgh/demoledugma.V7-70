import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.051_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "DeleteSheetIfExists" in line or "ws.Delete" in line or "Delete" in line:
        if "BuildReview" in line or True: # Actually let's just find all deletions
            pass
            
in_build_review = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview" in line:
        in_build_review = True
    if in_build_review and "Delete" in line:
        print(f"[{i}] {lines[i].strip()}")
    if in_build_review and "End Sub" in line:
        in_build_review = False
        break
