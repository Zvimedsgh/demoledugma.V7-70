with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

new_cleanup = """' ---- Remove ONLY the specific buttons we are about to recreate ----
2790 On Error Resume Next
Dim shapesToDelete As Variant
shapesToDelete = Array("btnBuildReview", "btnApplyCorrections", "btnShowResults", _
                       "btnBuildPresentation", "btnSaveReports", "btnViewReports", _
                       "btnNewClients", "btnSearchClient", "btnAllClients", _
                       "btnResetDefaults", "btnShowHidden", "btnNavExit", "shpBtn", "shpDemoMsgText")
Dim sName As Variant
For Each sName In shapesToDelete
    wsMain.Shapes(sName).Delete
Next sName
Err.Clear
2840 On Error GoTo ERR_HANDLER"""

# Replace the destructive cleanup block
text = re.sub(r"' ---- Remove ALL old buttons and stubborn floating shapes ----.*?2840 On Error GoTo ERR_HANDLER", new_cleanup, text, flags=re.DOTALL | re.IGNORECASE)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated shape cleanup!")
