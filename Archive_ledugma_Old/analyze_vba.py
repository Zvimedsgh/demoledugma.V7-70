import re

def analyze_vba(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    stack = []
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line or line.startswith("'"):
            continue
            
        # Ignore lines with line continuations if they are mid-statement
        # Simple heuristic:
        if line.lower().startswith("if ") and " then" in line.lower():
            if not line.lower().endswith(" then"):
                # One-liner If, ignore
                pass
            else:
                stack.append(("If", i + 1))
        elif line.lower() == "end if":
            if stack and stack[-1][0] == "If":
                stack.pop()
            else:
                print(f"Unmatched End If at line {i + 1}: {line}")
                
        elif line.lower().startswith("sub ") or line.lower().startswith("private sub ") or line.lower().startswith("public sub "):
            stack.append(("Sub", i + 1))
        elif line.lower() == "end sub":
            if stack and stack[-1][0] == "Sub":
                stack.pop()
            else:
                print(f"Unmatched End Sub at line {i + 1}: {line}")
                
        elif line.lower().startswith("function ") or line.lower().startswith("private function ") or line.lower().startswith("public function "):
            stack.append(("Function", i + 1))
        elif line.lower() == "end function":
            if stack and stack[-1][0] == "Function":
                stack.pop()
            else:
                print(f"Unmatched End Function at line {i + 1}: {line}")

    if stack:
        print("Unmatched blocks remaining:")
        for block, line_num in stack:
            print(f"  {block} started at line {line_num}")
    else:
        print("No mismatched basic blocks found.")

if __name__ == "__main__":
    analyze_vba(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas')
