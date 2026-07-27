import os
import shutil

src = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.039.bas'
dst = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.039_20260702_1445.bas'
shutil.copyfile(src, dst)
print("Copied successfully.")
