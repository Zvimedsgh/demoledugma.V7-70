import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.22.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Force exact button tops to guarantee perfect 44px spacing
    content = content.replace('wsMain.Range("C4").Top', 'wsMain.Range("C2").Top + 44')
    content = content.replace('wsMain.Range("C6").Top', 'wsMain.Range("C2").Top + 88')
    content = content.replace('wsMain.Range("C8").Top', 'wsMain.Range("C2").Top + 132')
    content = content.replace('wsMain.Range("C10").Top', 'wsMain.Range("C2").Top + 176')
    content = content.replace('wsMain.Range("C12").Top', 'wsMain.Range("C2").Top + 220')

    # 2. Fix A22 alignment to xlCenter so it looks better
    content = content.replace('wsMain.Range("A22").HorizontalAlignment = xlLeft', 'wsMain.Range("A22").HorizontalAlignment = xlCenter')

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
