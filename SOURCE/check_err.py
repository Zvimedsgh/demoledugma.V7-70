with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.59.bas', 'r', encoding='utf-8') as f:
    v359 = f.read()

idx1 = v359.find("' Final Summary Slide")
idx2 = v359.find("        ' Re-hide sheets that were unhidden")
print(v359[idx1:idx2].count('ERR_HANDLER:'))
