import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.055_20260702_1655.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    new_lines.append(line)

# Let's write a robust block fixer
def find_and_fix_blocks(lines):
    import re
    fixed_lines = list(lines)
    
    # We will do a full parse
    stack = []
    
    i = 0
    while i < len(fixed_lines):
        line = fixed_lines[i].strip()
        
        # Strip line numbers for checking
        m = re.match(r'^(\d+)\s+(.*)', line)
        if m:
            line_no_num = m.group(2)
        else:
            line_no_num = line
            
        if line_no_num.startswith("If ") and " Then " in line_no_num and not line_no_num.endswith(" Then") and not line_no_num.endswith("_"):
            pass
        elif line_no_num.startswith("If ") and (line_no_num.endswith(" Then") or line_no_num.endswith("_")):
            stack.append(("If", i))
        elif line_no_num.startswith("For "):
            stack.append(("For", i))
        elif line_no_num.startswith("Do While "):
            stack.append(("Do", i))
        elif line_no_num.startswith("With "):
            stack.append(("With", i))
        elif line_no_num.startswith("Select Case "):
            stack.append(("Select", i))
        elif line_no_num.startswith("Private Sub ") or line_no_num.startswith("Public Sub ") or line_no_num.startswith("Private Function ") or line_no_num.startswith("Public Function "):
            stack.append(("Sub", i))
        
        elif line_no_num.startswith("End If"):
            if len(stack) > 0 and stack[-1][0] == "If":
                stack.pop()
            else:
                print(f"Removed extra End If at {i+1}")
                fixed_lines[i] = "' " + fixed_lines[i]
                
        elif line_no_num.startswith("Next "):
            if len(stack) > 0 and stack[-1][0] == "For":
                stack.pop()
                
        elif line_no_num.startswith("Loop") or line_no_num == "Exit Do":
            if line_no_num.startswith("Loop"):
                if len(stack) > 0 and stack[-1][0] == "Do":
                    stack.pop()
                    
        elif line_no_num.startswith("End With"):
            if len(stack) > 0 and stack[-1][0] == "With":
                stack.pop()
                
        elif line_no_num.startswith("End Select"):
            if len(stack) > 0 and stack[-1][0] == "Select":
                stack.pop()
                
        elif line_no_num.startswith("End Sub") or line_no_num.startswith("End Function"):
            if len(stack) > 0 and stack[-1][0] == "Sub":
                stack.pop()
            else:
                # If stack is not empty, there is an unclosed block inside the sub!
                while len(stack) > 0 and stack[-1][0] != "Sub":
                    unclosed = stack.pop()
                    print(f"Unclosed {unclosed[0]} at {unclosed[1]+1} before End Sub/Function at {i+1}")
                    fixed_lines[i] = "End " + unclosed[0] + "\n" + fixed_lines[i]
                if len(stack) > 0 and stack[-1][0] == "Sub":
                    stack.pop()
        
        i += 1
        
    return fixed_lines

lines = find_and_fix_blocks(lines)

# Bump version to 2.056
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        lines[i] = line.replace("V2_055", "V2_056")
    if "Private Const APP_VERSION As String =" in line:
        lines[i] = line.replace("2.055", "2.056")
    if "VERSION: V2.055" in line:
        lines[i] = line.replace("2.055", "2.056")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.056.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Created V2.056")
