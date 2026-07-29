import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # The block we need to fix is shpDismissMsg text assignment
    # Pattern: .Text = ChrW(1492) ... followed by .Font.Size = 10 and .Font.Color
    
    correct_text = """3898         .Text = ChrW(1492) & ChrW(1489) & ChrW(1504) & ChrW(1514) & ChrW(1497) & ChrW(44) & ChrW(32) & ChrW(1488) & ChrW(1500) & ChrW(32) & ChrW(1514) & ChrW(1510) & ChrW(1497) & ChrW(1490) & ChrW(32) & ChrW(1497) & ChrW(1493) & ChrW(1514) & ChrW(1512) & ChrW(32) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & ChrW(32) & ChrW(1494) & ChrW(1492)"""

    # We use regex to find the section for shpDismissMsg and replace its text.
    pattern = re.compile(r'(3897     With shpDismiss\.TextFrame2\.TextRange\n).*?(    3900         \.Font\.Size = 10)', re.DOTALL)
    
    if pattern.search(content):
        replacement = r'\1' + correct_text + r'\n\2'
        new_content = pattern.sub(replacement, content)
        with open(bas_file, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print("Successfully fixed shpDismissMsg text!")
    else:
        print("Could not find the shpDismissMsg block.")

if __name__ == "__main__":
    main()
