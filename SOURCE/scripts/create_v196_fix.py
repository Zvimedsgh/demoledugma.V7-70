import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.195.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix BuildReview
old_rev = 'Public Sub BuildReview()\n    LogDebug "BuildReview STARTED"'
new_rev = 'Public Sub BuildReview()\n    InitRawColumns\n    LogDebug "BuildReview STARTED"'
content = content.replace(old_rev, new_rev)

# Fix ApplyCorrections
old_app = 'Public Sub ApplyCorrectionsAndBuildReports()\n    LogDebug "ApplyCorrectionsAndBuildReports STARTED"'
new_app = 'Public Sub ApplyCorrectionsAndBuildReports()\n    InitRawColumns\n    LogDebug "ApplyCorrectionsAndBuildReports STARTED"'
content = content.replace(old_app, new_app)

# Update version strings
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_195"', 'Attribute VB_Name = "Goren_Claude_V2_196"')
content = content.replace('VERSION: V2.195', 'VERSION: V2.196')
content = content.replace('APP_VERSION As String = "2.195"', 'APP_VERSION As String = "2.196"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.196")
