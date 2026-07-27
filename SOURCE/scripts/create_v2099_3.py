import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.095.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.099.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_build_review = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview()" in line:
        in_build_review = True
    
    if in_build_review and "If lastRow < 2 Then Err.Raise" in line:
        new_lines.append(line)
        new_lines.append("    ' Populate srcData array for ultra-fast processing\n")
        new_lines.append("    srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, 100)).Value2\n")
        in_build_review = False # We found it and inserted it, don't insert again
        continue
        
    new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.099 created.")
