import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.208.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Public Sub ApplyCorrectionsAndBuildReports" in line:
        # Find ERR_HANDLER below this
        for j in range(i, len(lines)):
            if "ERR_HANDLER:" in lines[j]:
                for k in range(j, j+15):
                    print(f"[{k+1}] {lines[k].strip()}")
                break
        break
