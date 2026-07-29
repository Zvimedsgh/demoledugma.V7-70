import re

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.54.bas', 'r', encoding='utf-8') as f:
    v354 = f.read()

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.56.bas', 'r', encoding='utf-8') as f:
    v356 = f.read()

# Update VB_Name
v357 = v356.replace('Attribute VB_Name = "Goren_Claude_V3_56"', 'Attribute VB_Name = "Goren_Claude_V3_57"')

# Extract the missing cleanup logic from V3.54
start_marker = "        Dim pg As Long"
end_marker = "' Clean up presentation sheets automatically"

idx1 = v354.find(start_marker)
idx2 = v354.find(end_marker)

if idx1 != -1 and idx2 != -1:
    missing_logic = v354[idx1:idx2]
else:
    print("Failed to find missing logic in V3.54")
    exit(1)

# In V3.57 (which was copied from V3.56), we need to insert this missing logic
# Right after the end of the coffee slide logic.
# The coffee slide logic ends with:
#         btnBack.ActionSettings(1).SoundEffect.Name = "Click"
#         ' ===============================

insert_marker = "        ' ===============================\n"
idx3 = v357.find(insert_marker, v357.find("btnBack.ActionSettings"))

if idx3 != -1:
    # Insert the missing logic right after the marker
    v357 = v357[:idx3 + len(insert_marker)] + "\n" + missing_logic + v357[idx3 + len(insert_marker):]
else:
    print("Failed to find insertion point in V3.57")
    exit(1)

# Now, why did AddPicture fail?
# In PowerPoint, AddPicture requires Single for Left, Top, Width, Height.
# And we used 0, -1 for LinkToFile and SaveWithDocument. But sometimes these Enums need to be msoFalse and msoTrue implicitly.
# Let's change the AddPicture line to not use 0, -1 if they are failing, but they usually work.
# Actually, maybe the picture was just inserted too large or small? 500x500 is fine.
# Let's use `msoFalse` and `msoTrue` explicitly if possible, but VBA needs them as constants.
# We will just write 0, -1.
# What if we just use an online image url for the coffee if the local file is failing?
# Wait! Let's check the AddPicture line.
v357 = v357.replace(
    'Set picCoffee = sldCoffee.Shapes.AddPicture(coffeePath, 0, -1, 230, 20, 500, 500)',
    'Set picCoffee = sldCoffee.Shapes.AddPicture(coffeePath, 0, -1, 230, 20, 500, 500)'
)

with open(r'C:\LEVAV PROJECT\SOURCE\Goren_Claude_V3.57.bas', 'w', encoding='utf-8') as f:
    f.write(v357)

print("V3.57 created successfully.")
