import os
import glob
import shutil

source_dir = r"C:\LEVAV PROJECT\SOURCE"
root_dir = r"C:\LEVAV PROJECT"
archive_dir = r"C:\LEVAV PROJECT\Archive_V3.70"

if not os.path.exists(archive_dir):
    os.makedirs(archive_dir)

files_to_move = []

# Collect old .bas files
for bas_file in glob.glob(os.path.join(source_dir, "Goren_Claude_*.bas")):
    if not bas_file.endswith("Goren_claude_Orit_final_V3.70.bas"):
        files_to_move.append(bas_file)

# Collect python scripts in SOURCE
for py_file in glob.glob(os.path.join(source_dir, "*.py")):
    files_to_move.append(py_file)

# Collect specific txt files in SOURCE
txt_files = ["old_bp.txt", "old_bp_clean.txt", "new_bp_fixed.txt", "v351_bp.txt", "print_deleted.py"]
for f in txt_files:
    path = os.path.join(source_dir, f)
    if os.path.exists(path):
        files_to_move.append(path)

# Collect coffee images in SOURCE
for ext in ["*.jpg", "*.png"]:
    for f in glob.glob(os.path.join(source_dir, "coffee" + ext)):
        files_to_move.append(f)

# Collect python, vbs, png files in ROOT
for py_file in glob.glob(os.path.join(root_dir, "*.py")):
    files_to_move.append(py_file)
for vbs_file in glob.glob(os.path.join(root_dir, "*.vbs")):
    if "install_modLevav.vbs" not in vbs_file:  # keep the original installer
        files_to_move.append(vbs_file)
for png_file in glob.glob(os.path.join(root_dir, "test_*.png")):
    files_to_move.append(png_file)

# Move the files
moved_count = 0
for f_path in set(files_to_move):
    try:
        if os.path.isfile(f_path):
            basename = os.path.basename(f_path)
            shutil.move(f_path, os.path.join(archive_dir, basename))
            moved_count += 1
            print(f"Moved {basename}")
    except Exception as e:
        print(f"Failed to move {f_path}: {e}")

print(f"Total files moved: {moved_count}")
