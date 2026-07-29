with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.50.bas', 'r', encoding='utf-8') as f:
    v350 = f.read()

with open(r'C:\LEVAV PROJECT\SOURCE\old_bp_clean.txt', 'r', encoding='utf-8') as f:
    bp_clean = f.read()

start = v350.find('Public Sub BuildPresentation()')
end = v350.find('End Sub', start) + 7

print('Length of BuildPresentation in V3.50:', end - start)
print('Length of bp_clean:', len(bp_clean))

v_final = v350[:start] + bp_clean + v350[end:]

# Set VB_Name and Version
v_final = v_final.replace('Attribute VB_Name = "Goren_Claude_V3_50"', 'Attribute VB_Name = "Goren_Claude_Final_System"')
v_final = v_final.replace("' VERSION: V3.50", "' VERSION: Final_System\n' CHANGES IN Final_System: Re-integrated all helper functions correctly.")

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Final_System.bas', 'w', encoding='utf-8') as f:
    f.write(v_final)

print('Created Final System')
