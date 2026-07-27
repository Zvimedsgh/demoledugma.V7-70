import re

def main():
    filepath = r"c:\LEVAV PROJECT\SOURCE\Goren_Claude_V1.28.bas"
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. SetupMainSheet G7 default
    old_setup_g7 = 'wsMain.Range("G7").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dateTypeList\n    wsMain.Range("G8").Validation.Delete'
    new_setup_g7 = 'wsMain.Range("G7").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dateTypeList\n    wsMain.Range("G7").Value = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)\n    wsMain.Range("G8").Validation.Delete'
    content = content.replace(old_setup_g7, new_setup_g7)

    # 2. ResetHomeDefaults G7 default
    old_reset_home = 'On Error GoTo 0\n    \' G9 ='
    new_reset_home = 'On Error GoTo 0\n    wsMain.Range("G7").Value = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)\n    \' G9 ='
    content = content.replace(old_reset_home, new_reset_home)

    # Bump version
    content = content.replace('Attribute VB_Name = "Goren_Claude1.28"', 'Attribute VB_Name = "Goren_Claude1.29"')
    content = content.replace("' VERSION: V1.28", "' VERSION: V1.29")
    content = content.replace('Private Const APP_VERSION As String = "1.28"', 'Private Const APP_VERSION As String = "1.29"')
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Tweaks applied successfully.")

if __name__ == "__main__":
    main()
