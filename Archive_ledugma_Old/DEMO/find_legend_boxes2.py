from PIL import Image

img = Image.open(r"C:\ledugma\DEMO\page13.png")
pixels = img.load()

# Look for orange and green on the FAR RIGHT (X > 1700 out of 2000, which is X > 850 in PDF units)
orange_pixels = []
green_pixels = []
for y in range(img.height):
    for x in range(int(img.width * 0.85), img.width):
        r, g, b = pixels[x, y][:3]
        if 220 < r < 255 and 100 < g < 140 and 20 < b < 60:
            orange_pixels.append((x, y))
        if 90 < r < 130 and 150 < g < 190 and 50 < b < 90:
            green_pixels.append((x, y))

if orange_pixels:
    ox = [p[0] for p in orange_pixels]
    oy = [p[1] for p in orange_pixels]
    print(f"Orange legend roughly at X: {min(ox)} to {max(ox)}, Y: {min(oy)} to {max(oy)}")

if green_pixels:
    gx = [p[0] for p in green_pixels]
    gy = [p[1] for p in green_pixels]
    print(f"Green legend roughly at X: {min(gx)} to {max(gx)}, Y: {min(gy)} to {max(gy)}")
