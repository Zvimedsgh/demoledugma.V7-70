import re

file_path = r"C:\ledugma\DEMO\הוראות_שימוש_V9.32.html"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Update version from V7.70 to V9.32
content = content.replace("Demo_Reports_Syatem_V7.70.xlsm", "Demo_Reports_System_V9.32.xlsm")
content = content.replace("Demo_Reports_System_V7.70.xlsm", "Demo_Reports_System_V9.32.xlsm")

# We can also add a line about 2024/2025 being display-only
new_instruction = r"<li><strong>בחירת שנת בסיס ושנה נוכחית:</strong> בהדגמה אין לשנות את השנים (רק נתוני 2019/2020 זמינים ולכן המערכת מציגה 2024/2025 לשם ההדגמה בלבד, אך תעבד את הנתונים ההיסטוריים). <i>במוצר המוגמר מפיקים נתונים מתוכנת ניהול הביטוחים עבור שנת הבסיס והנוכחית ושומרים אותם במיקום מוגדר מראש על המחשב.</i></li>"
old_instruction = r"<li><strong>בחירת שנת בסיס ושנה נוכחית:</strong> בהדגמה אין לשנות את השנים (רק נתוני 2019/2020 זמינים). <i>במוצר המוגמר מפיקים נתונים מתוכנת ניהול הביטוחים עבור שנת הבסיס והנוכחית ושומרים אותם במיקום מוגדר מראש על המחשב.</i></li>"

content = content.replace(old_instruction, new_instruction)

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Updated HTML")
