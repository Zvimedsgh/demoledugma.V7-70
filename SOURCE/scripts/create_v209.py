import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.208.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.209.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix SendForReview tempPath
old_sfr = """tempPath = Environ$("TEMP") & "\\Exceptions_" & Format$(Date, "yyyy-mm-dd") & ".xlsx\""""
new_sfr = """tempPath = Environ$("TEMP") & "\\Exceptions_" & Format$(Now, "yyyy-mm-dd_hhmmss") & ".xlsx\""""
content = content.replace(old_sfr, new_sfr)

# Fix EmailNewClients tempPath
old_enc = """tempPath = Environ$("TEMP") & "\\" & outSheetName & "_" & Format$(Date, "yyyy-mm-dd") & ".xlsx\""""
new_enc = """tempPath = Environ$("TEMP") & "\\" & outSheetName & "_" & Format$(Now, "yyyy-mm-dd_hhmmss") & ".xlsx\""""
content = content.replace(old_enc, new_enc)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_208"', 'Attribute VB_Name = "Goren_Claude_V2_209"')
content = content.replace('VERSION: V2.208', 'VERSION: V2.209')
content = content.replace('APP_VERSION As String = "2.208"', 'APP_VERSION As String = "2.209"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.209")
