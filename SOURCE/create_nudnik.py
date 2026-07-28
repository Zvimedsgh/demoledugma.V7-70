with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_System_Final.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('Attribute VB_Name = "Goren_Claude_V3_62"', 'Attribute VB_Name = "Goren_Claude_Nudnik_Benleumi"')
text = text.replace("' VERSION: V3.62\n' CHANGES IN V3.62: Final robust version.", "' VERSION: Nudnik Benleumi\n' CHANGES IN Nudnik Benleumi: Undeniably the right file.")

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Nudnik_Benleumi.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print('Nudnik Benleumi created')
