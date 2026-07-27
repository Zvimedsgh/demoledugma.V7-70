import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.194.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace Constants with Variables
old_consts = """Private Const RAW_CUSTOMER As Long = 1
Private Const RAW_CUSTNAME As Long = 2
Private Const RAW_POLICY As Long = 11
Private Const RAW_ADDENDUM As Long = 12
Private Const RAW_COMPNUM As Long = 13
Private Const RAW_COMPANY As Long = 14
Private Const RAW_BRANCHNUM As Long = 15
Private Const RAW_BRANCHNAME As Long = 16
Private Const RAW_INSURANCE_START As Long = 17
Private Const RAW_BORDEREU As Long = 19
Private Const RAW_AGENTNUM As Long = 20
Private Const RAW_AGENTNAME As Long = 21
Private Const RAW_TELLERNUM As Long = 24
Private Const RAW_TELLERNAME As Long = 25
Private Const RAW_PREMIUM As Long = 28
Private Const RAW_COMMISSION As Long = 32
Private Const RAW_CURRENCY As Long = 27
Private Const RAW_ACTIONCOL As Long = 39
Private Const RAW_IDNUMBER As Long = 45"""

new_vars = """Private RAW_CUSTOMER As Long
Private RAW_CUSTNAME As Long
Private RAW_POLICY As Long
Private RAW_ADDENDUM As Long
Private RAW_COMPNUM As Long
Private RAW_COMPANY As Long
Private RAW_BRANCHNUM As Long
Private RAW_BRANCHNAME As Long
Private RAW_INSURANCE_START As Long
Private RAW_BORDEREU As Long
Private RAW_AGENTNUM As Long
Private RAW_AGENTNAME As Long
Private RAW_TELLERNUM As Long
Private RAW_TELLERNAME As Long
Private RAW_PREMIUM As Long
Private RAW_COMMISSION As Long
Private RAW_CURRENCY As Long
Private RAW_ACTIONCOL As Long
Private RAW_IDNUMBER As Long

Private Sub InitRawColumns()
    Dim dictFieldCol As Object, dictFieldDisp As Object
    Set dictFieldCol = CreateObject("Scripting.Dictionary")
    Set dictFieldDisp = CreateObject("Scripting.Dictionary")
    dictFieldCol.CompareMode = vbTextCompare
    dictFieldDisp.CompareMode = vbTextCompare
    
    LoadCheckedFields ThisWorkbook.Worksheets(H_SET_FIELDMAP()), dictFieldCol, dictFieldDisp
    
    ' Set defaults first, just in case they are missing from the sheet
    RAW_CUSTOMER = 1
    RAW_CUSTNAME = 2
    RAW_POLICY = 11
    RAW_ADDENDUM = 12
    RAW_COMPNUM = 13
    RAW_COMPANY = 14
    RAW_BRANCHNUM = 15
    RAW_BRANCHNAME = 16
    RAW_INSURANCE_START = 17
    RAW_BORDEREU = 19
    RAW_AGENTNUM = 20
    RAW_AGENTNAME = 21
    RAW_TELLERNUM = 24
    RAW_TELLERNAME = 25
    RAW_CURRENCY = 27
    RAW_PREMIUM = 28
    RAW_COMMISSION = 32
    RAW_ACTIONCOL = 39
    RAW_IDNUMBER = 45
    
    ' Override with dynamic mapping
    If dictFieldCol.Exists("CUSTOMER_NUMBER") Then RAW_CUSTOMER = dictFieldCol("CUSTOMER_NUMBER")
    If dictFieldCol.Exists("CUSTOMER_NAME") Then RAW_CUSTNAME = dictFieldCol("CUSTOMER_NAME")
    If dictFieldCol.Exists("POLICY") Then RAW_POLICY = dictFieldCol("POLICY")
    If dictFieldCol.Exists("ADDENDUM") Then RAW_ADDENDUM = dictFieldCol("ADDENDUM")
    If dictFieldCol.Exists("COMPANY_NUMBER") Then RAW_COMPNUM = dictFieldCol("COMPANY_NUMBER")
    If dictFieldCol.Exists("COMPANY_NAME") Then RAW_COMPANY = dictFieldCol("COMPANY_NAME")
    If dictFieldCol.Exists("BRANCH_NUMBER") Then RAW_BRANCHNUM = dictFieldCol("BRANCH_NUMBER")
    If dictFieldCol.Exists("BRANCH_NAME") Then RAW_BRANCHNAME = dictFieldCol("BRANCH_NAME")
    If dictFieldCol.Exists("INSURANCE_START_DATE") Then RAW_INSURANCE_START = dictFieldCol("INSURANCE_START_DATE")
    If dictFieldCol.Exists("BORDEREAU_DATE") Then RAW_BORDEREU = dictFieldCol("BORDEREAU_DATE")
    If dictFieldCol.Exists("AGENT_NUMBER") Then RAW_AGENTNUM = dictFieldCol("AGENT_NUMBER")
    If dictFieldCol.Exists("AGENT_NAME") Then RAW_AGENTNAME = dictFieldCol("AGENT_NAME")
    If dictFieldCol.Exists("UNDERWRITER_TELLER_NUMBER") Then RAW_TELLERNUM = dictFieldCol("UNDERWRITER_TELLER_NUMBER")
    If dictFieldCol.Exists("UNDERWRITER_TELLER_NAME") Then RAW_TELLERNAME = dictFieldCol("UNDERWRITER_TELLER_NAME")
    If dictFieldCol.Exists("CURRENCY") Then RAW_CURRENCY = dictFieldCol("CURRENCY")
    If dictFieldCol.Exists("PREMIUM") Then RAW_PREMIUM = dictFieldCol("PREMIUM")
    If dictFieldCol.Exists("COMPANY_COMMISSION") Then RAW_COMMISSION = dictFieldCol("COMPANY_COMMISSION")
    If dictFieldCol.Exists("ACTION") Then RAW_ACTIONCOL = dictFieldCol("ACTION")
    If dictFieldCol.Exists("ID_NUMBER") Then RAW_IDNUMBER = dictFieldCol("ID_NUMBER")
End Sub"""

content = content.replace(old_consts, new_vars)

# Call InitRawColumns at the start of ApplyCorrectionsAndBuildReports
app_old = """Public Sub ApplyCorrectionsAndBuildReports()
    Application.ScreenUpdating = False
    Application.EnableEvents = False"""

app_new = """Public Sub ApplyCorrectionsAndBuildReports()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call InitRawColumns"""
content = content.replace(app_old, app_new)

# Call InitRawColumns at the start of BuildReview
rev_old = """Public Sub BuildReview()
    Application.ScreenUpdating = False
    Application.EnableEvents = False"""

rev_new = """Public Sub BuildReview()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Call InitRawColumns"""
content = content.replace(rev_old, rev_new)

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_193"', 'Attribute VB_Name = "Goren_Claude_V2_194"')
content = content.replace('VERSION: V2.193', 'VERSION: V2.194')
content = content.replace('APP_VERSION As String = "2.193"', 'APP_VERSION As String = "2.194"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.194")
