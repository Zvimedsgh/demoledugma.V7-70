import re

def main():
    bas221 = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.221.bas'
    bas222 = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    
    with open(bas221, 'r', encoding='utf-8') as f:
        content221 = f.read()
        
    pattern = re.compile(r"(' ============================================================================\n' HELPER: Apply zebra striping.*?End Sub\n)", re.DOTALL)
    match = pattern.search(content221)
    
    if match:
        zebra_code = match.group(1)
        
        with open(bas222, 'r', encoding='utf-8') as f:
            content222 = f.read()
            
        # Insert before ApplyDemoLockOnOpen
        insert_point = "Public Sub ApplyDemoLockOnOpen("
        if insert_point in content222:
            content222 = content222.replace(insert_point, zebra_code + "\n\n" + insert_point)
            
            with open(bas222, 'w', encoding='utf-8') as f:
                f.write(content222)
            print("Successfully restored ApplyZebraStriping to V2.222.bas")
        else:
            print("Could not find insertion point in V2.222.bas")
    else:
        print("Could not find ApplyZebraStriping in V2.221.bas")

if __name__ == "__main__":
    main()
