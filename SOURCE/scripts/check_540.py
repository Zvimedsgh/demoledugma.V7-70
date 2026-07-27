import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.206.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# find "540" inside BuildReview
for i, line in enumerate(lines):
    if "Sub BuildReview" in line:
        for j in range(i, i+300):
            if line.strip().startswith("540"):
                pass # not helpful to just look for "540" since lines changed
            
        break

for i, line in enumerate(lines):
    if "SOURCE FILE NOT FOUND FOR YEAR" in line:
        for j in range(i-5, i+5):
            print(f"[{j+1}] {lines[j].strip()}")

