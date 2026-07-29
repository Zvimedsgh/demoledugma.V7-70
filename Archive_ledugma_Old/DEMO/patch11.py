import re

file_path = r"C:\ledugma\DEMO\הוראות_שימוש_V9.32.html"

with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Replace the overly transparent instruction with the more professional one
old_instruction = r"בהדגמה אין לשנות את השנים (רק נתוני 2019/2020 זמינים ולכן המערכת מציגה 2024/2025 לשם ההדגמה בלבד, אך תעבד את הנתונים ההיסטוריים)."
new_instruction = r"בהדגמה לא ניתן לשנות את השנים הנבדקות, הדבר אפשרי רק בגירסה המלאה."

content = content.replace(old_instruction, new_instruction)

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Updated HTML with professional text")
