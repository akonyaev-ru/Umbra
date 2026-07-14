import os
from PIL import Image, ImageEnhance

img_path = "logo.png"
if not os.path.exists(img_path):
    print("File not found:", img_path)
    exit(1)

img = Image.open(img_path)
# Ensure the image is RGBA
img = img.convert("RGBA")

sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (24, 24), (16, 16)]
imgs = []

for s in sizes:
    r = img.resize(s, Image.Resampling.LANCZOS)
    if s[0] <= 64:
        # Separate alpha channel to avoid sharpening edges becoming dark/weird
        rgb = r.convert("RGB")
        alpha = r.split()[3]
        
        # Apply sharpness only to RGB
        enhancer = ImageEnhance.Sharpness(rgb)
        rgb_sharp = enhancer.enhance(3.5) # Strong sharpening for crispness
        
        # Recombine
        r = Image.merge("RGBA", (*rgb_sharp.split(), alpha))
    imgs.append(r)

imgs[0].save('icon_purple.ico', format='ICO', sizes=sizes, append_images=imgs[1:])
print("Saved super-crisp icon_purple.ico")
