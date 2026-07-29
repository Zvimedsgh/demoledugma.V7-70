import os
import glob
import shutil

root_dir = r"c:\ledugma"
archive_dir = os.path.join(root_dir, "Archive_Legacy")

if not os.path.exists(archive_dir):
    os.makedirs(archive_dir)

moved_count = 0
# Get all files and directories in root_dir
for item in os.listdir(root_dir):
    item_path = os.path.join(root_dir, item)
    # Skip the archive directory itself and .git folder
    if item == "Archive_Legacy" or item == ".git":
        continue
    
    # We move both files and subdirectories if any, but let's stick to files for safety
    if os.path.isfile(item_path):
        try:
            shutil.move(item_path, os.path.join(archive_dir, item))
            moved_count += 1
            print(f"Moved {item}")
        except Exception as e:
            print(f"Failed to move {item}: {e}")

print(f"Total files moved: {moved_count}")
