import re
import datetime

def create_v971():
    with open(r'C:\ledugma\DEMO\V9.70_Final.txt', 'r', encoding='windows-1255', errors='ignore') as f:
        code = f.read()

    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M")
    
    # Update version
    code = code.replace("9.70", "9.71")
    code = code.replace("V9.70", "V9.71")

    # Quarter match: change to "רבע" (1512, 1489, 1506)
    # Old: InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare)
    # New: InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506), vbTextCompare)
    code = re.sub(
        r'InStr\(1, periodType, ChrW\(1512\) & ChrW\(1489\) & ChrW\(1506\) & ChrW\(1493\) & ChrW\(1504\) & ChrW\(1497\), vbTextCompare\)',
        r'InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506), vbTextCompare)',
        code
    )

    # Month match: change to "חודש" (1495, 1493, 1491, 1513)
    # Old: InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare)
    # New: InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513), vbTextCompare)
    code = re.sub(
        r'InStr\(1, periodType, ChrW\(1495\) & ChrW\(1493\) & ChrW\(1491\) & ChrW\(1513\) & ChrW\(1497\), vbTextCompare\)',
        r'InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513), vbTextCompare)',
        code
    )

    # Year match: change to "שנ" (1513, 1504)
    # Old: InStr(1, periodType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare)
    # New: InStr(1, periodType, ChrW(1513) & ChrW(1504), vbTextCompare)
    code = re.sub(
        r'InStr\(1, periodType, ChrW\(1513\) & ChrW\(1504\) & ChrW\(1514\) & ChrW\(1497\), vbTextCompare\)',
        r'InStr(1, periodType, ChrW(1513) & ChrW(1504), vbTextCompare)',
        code
    )

    with open(r'C:\ledugma\DEMO\V9.71_Final.txt', 'w', encoding='windows-1255') as f:
        f.write(code)
        
    print("V9.71_Final.txt created successfully.")

if __name__ == '__main__':
    create_v971()
