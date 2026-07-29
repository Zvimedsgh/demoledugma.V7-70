from PIL import Image

img = Image.open(r"C:\ledugma\DEMO\page13.png")
pixels = img.load()

# The image is 960 * (150/72) = 2000 pixels wide
# The squares are orange and green.
# Let's find orange (around 237, 125, 49) and green (around 112, 173, 71)

orange_pixels = []
green_pixels = []
for y in range(img.height):
    for x in range(img.width // 2, img.width): # only right half
        r, g, b = pixels[x, y][:3]
        if 220 < r < 255 and 100 < g < 140 and 20 < b < 60:
            orange_pixels.append((x, y))
        if 90 < r < 130 and 150 < g < 190 and 50 < b < 90:
            green_pixels.append((x, y))

if orange_pixels:
    ox = [p[0] for p in orange_pixels]
    oy = [p[1] for p in orange_pixels]
    print(f"Orange box roughly at X: {min(ox)} to {max(ox)}, Y: {min(oy)} to {max(oy)}")

if green_pixels:
    gx = [p[0] for p in green_pixels]
    gy = [p[1] for p in green_pixels]
    print(f"Green box roughly at X: {min(gx)} to {max(gx)}, Y: {min(gy)} to {max(gy)}")
