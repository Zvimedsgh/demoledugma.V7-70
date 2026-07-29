import zipfile
import os

xlsm_path = r"C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70.xlsm"
temp_dir = r"C:\ledugma\DEMO\xlsm_temp"

# Extract the xlsm
with zipfile.ZipFile(xlsm_path, 'r') as zip_ref:
    zip_ref.extractall(temp_dir)

vba_path = os.path.join(temp_dir, "xl", "vbaProject.bin")

# Read binary
with open(vba_path, "rb") as f:
    data = f.read()

# Replace ascii strings
data = data.replace(b"2019", b"2024")
data = data.replace(b"2020", b"2025")

# Replace utf-16le strings
data = data.replace("2019".encode("utf-16le"), "2024".encode("utf-16le"))
data = data.replace("2020".encode("utf-16le"), "2025".encode("utf-16le"))

# Write binary
with open(vba_path, "wb") as f:
    f.write(data)

# Re-zip
patched_xlsm = r"C:\ledugma\DEMO\Demo_Reports_Syatem_V7.70_patched.xlsm"
with zipfile.ZipFile(patched_xlsm, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
    for root, dirs, files in os.walk(temp_dir):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, temp_dir)
            zip_ref.write(file_path, arcname)

print("Patched xlsm created.")
