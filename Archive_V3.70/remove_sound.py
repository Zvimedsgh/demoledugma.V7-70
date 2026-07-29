with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Super_Perfect.bas', 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    if 'SoundEffect.Name' not in line:
        new_lines.append(line)

text = ''.join(new_lines)
text = text.replace('Attribute VB_Name = "Goren_Claude_Super_Perfect"', 'Attribute VB_Name = "Goren_Claude_Orit_Final"')

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Orit_Final.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Orit Final created')
