import sys
import win32com.client

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.08.xlsm'

excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(filepath, ReadOnly=True)
for sh in wb.Sheets:
    print(sh.Name)
wb.Close(False)
excel.Quit()

