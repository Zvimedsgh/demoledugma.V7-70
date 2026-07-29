import re

def main():
    bas_file = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
    with open(bas_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # The block we need to fix is:
    # 3770 shp.Name = "btnShowHidden"
    # 3780 shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
    # 3790 shp.TextFrame2.TextRange.Text = ...
    #         .Font.Size = 10
    
    # Let's extract the exact original text assignment for btnShowHidden
    correct_text = """3790 shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & "/" & ChrW(1492) & ChrW(1505) & _
        ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & _
        ChrW(1514)
    3800 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    3810 shp.TextFrame2.TextRange.Font.Size = 10"""

    # We will use regex to find the section for btnShowHidden and replace its text.
    # Pattern: 3780 shp.Fill.ForeColor.RGB = RGB(80, 80, 80) ... 3820 shp.TextFrame2.TextRange.Font.Bold = msoTrue
    
    pattern = re.compile(r'(3780 shp\.Fill\.ForeColor\.RGB = RGB\(80, 80, 80\)\n).*?(    3820 shp\.TextFrame2\.TextRange\.Font\.Bold = msoTrue)', re.DOTALL)
    
    if pattern.search(content):
        replacement = r'\1    ' + correct_text + r'\n\2'
        new_content = pattern.sub(replacement, content)
        with open(bas_file, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print("Successfully fixed btnShowHidden text!")
    else:
        print("Could not find the btnShowHidden block.")

if __name__ == "__main__":
    main()
