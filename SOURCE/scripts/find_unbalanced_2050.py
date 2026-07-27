import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.050_20260702_1645.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

def check_unbalanced(lines):
    if_count = 0
    end_if_count = 0
    for i, line in enumerate(lines):
        line = line.strip()
        
        # Single line if
        if line.startswith("If ") and " Then " in line and not line.endswith(" Then") and not line.endswith("_"):
            continue
            
        if line.startswith("If ") and (line.endswith(" Then") or line.endswith("_")):
            if_count += 1
            
        if line.startswith("End If"):
            end_if_count += 1
            if end_if_count > if_count:
                print(f"Unbalanced End If at line {i+1}: {line}")
                return

check_unbalanced(lines)
print("Done checking")
