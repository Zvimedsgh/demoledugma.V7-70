with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.50.bas', 'r', encoding='utf-8') as f:
    v350 = f.read()

start = v350.find("' HELPER: Export Total Summary chart to image file")
end = v350.find("End Sub", start) + 7

export_total_chart = v350[start:end]
print('Found function:', len(export_total_chart), 'bytes')

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Nudnik_Benleumi.bas', 'r', encoding='utf-8') as f:
    text = f.read()

text = text + '\n\n' + export_total_chart + '\n'

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_Nudnik_Benleumi.bas', 'w', encoding='utf-8') as f:
    f.write(text)

print('Appended to Nudnik Benleumi')
