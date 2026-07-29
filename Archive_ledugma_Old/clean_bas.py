import re

file_path = r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'
with open(file_path, 'r', encoding='utf-8') as f:
    text = f.read()

# Replace case-sensitive to maintain CamelCase for variables
text = text.replace('LevavTemp_', 'DemoTemp_')
text = text.replace('modLevav', 'modDemo')
text = text.replace('Levav System', 'Goren System')
text = text.replace('levav_total', 'demo_total')
text = text.replace('levav_prem', 'demo_prem')
text = text.replace('levav_comm', 'demo_comm')
text = text.replace('levav_docs', 'demo_docs')
text = text.replace('levav_ins', 'demo_ins')
text = text.replace('imgNoLevav', 'imgNoDemo')
text = text.replace('levav_nolev', 'demo_nodem')
text = text.replace('noLevavOK', 'noDemoOK')
text = text.replace('levavName', 'demoName')
text = text.replace('Agents without Levav', 'Agents without Demo')
text = text.replace('agents without Levav', 'agents without Demo')
text = text.replace('lelo levav', 'lelo demo')

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(text)

print("Replaced all sensitive names in VBA code.")
