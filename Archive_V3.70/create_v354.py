import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.53.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update VB_Name
content = content.replace('Attribute VB_Name = "Goren_Claude_V3_53"', 'Attribute VB_Name = "Goren_Claude_V3_54"')

# We will comment out the offending ScreenUpdating=True lines inside the newly added subs.
# We know the text is exactly:
# Application.DisplayAlerts = True
# Application.ScreenUpdating = True
# Exit Sub

content = content.replace(
    'Application.DisplayAlerts = True\n    Application.ScreenUpdating = True\n    Exit Sub',
    'Application.DisplayAlerts = True\n    \'Application.ScreenUpdating = True\n    Exit Sub'
)

content = content.replace(
    'Application.DisplayAlerts = True\n    Application.ScreenUpdating = True\n    Err.Raise',
    'Application.DisplayAlerts = True\n    \'Application.ScreenUpdating = True\n    Err.Raise'
)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.54.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V3.54 created successfully.")
