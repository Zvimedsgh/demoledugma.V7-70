import os

folder = r"c:\LEVAV PROJECT\SOURCE\Data"
if os.path.exists(folder):
    files = os.listdir(folder)
    print("Files in Data folder:")
    for f in files:
        print(f)
else:
    print("Data folder does not exist!")

