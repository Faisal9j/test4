"""Brand assets for the landscape deck, cropped from the existing Amana presentation (28 slide images, 2560x1810)."""
import sys, os
from PIL import Image, ImageDraw
src, out = sys.argv[1], sys.argv[2]        # src = folder with image-N-1.jpg
os.makedirs(out, exist_ok=True)
S = 2560 / 937                              # tile units -> pixels
def L(n): return Image.open(f"{src}/image-{n}-1.jpg").convert("RGB")
def crop(n, x0, y0, x1, y1): return L(n).crop((int(x0 * S), int(y0 * S), int(x1 * S), int(y1 * S)))

def slim(img, path, width):
    img = img.resize((width, int(img.height * width / img.width)), Image.LANCZOS)
    img.quantize(256, method=Image.Quantize.FASTOCTREE, dither=Image.Dither.FLOYDSTEINBERG).save(path, optimize=True)

def rounded(img, path, width, r=0.06, corners=(1, 1, 1, 1)):
    img = img.convert("RGBA"); W, H = img.size; rr = int(min(W, H) * r)
    m = Image.new("L", (W * 2, H * 2), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, W * 2 - 1, H * 2 - 1), rr * 2, fill=255, corners=tuple(bool(c) for c in corners))
    img.putalpha(m.resize((W, H), Image.LANCZOS)); slim(img, path, width)

def fit(img, aspect):
    W, H = img.size
    if W / H > aspect: w = int(H * aspect); return img.crop(((W - w) // 2, 0, (W - w) // 2 + w, H))
    h = int(W / aspect); return img.crop((0, (H - h) // 2, W, (H - h) // 2 + h))

# portrait photos for divider / intro pages (aspect ~ 0.89)
P = {"tower": crop(2, 42, 100, 400, 585), "hotel": crop(9, 66, 116, 463, 563), "plaza": crop(13, 66, 116, 463, 563),
     "towers": crop(16, 66, 116, 463, 563), "building": crop(22, 66, 116, 463, 563), "night": crop(26, 42, 95, 395, 595)}
for k, im in P.items(): rounded(fit(im, 0.89), f"{out}/ph-{k}.png", 900, 0.12, (0, 1, 0, 1))
# cover: whole photo of slide 1 (with the logo notch); the baked title box is repainted so new text can go on it
cv = crop(1, 30, 30, 908, 585).convert("RGBA"); W, H = cv.size
bx0, by0 = int((308 - 30) * S), int((360 - 30) * S)
m = Image.new("L", (W * 2, H * 2), 0)
ImageDraw.Draw(m).rounded_rectangle((bx0 * 2, by0 * 2, W * 2 + 400, H * 2 + 400), 250, fill=255)
box = Image.new("RGBA", (W * 2, H * 2), (0, 106, 70, 255)); cv2 = cv.resize((W * 2, H * 2)); cv2.paste(box, (0, 0), m)
cv = cv2.resize((W, H), Image.LANCZOS)
rm = Image.new("L", (W * 2, H * 2), 0); ImageDraw.Draw(rm).rounded_rectangle((0, 0, W * 2 - 1, H * 2 - 1), 110, fill=255)
cv.putalpha(rm.resize((W, H), Image.LANCZOS)); slim(cv, f"{out}/cover.png", 2000)
# logo: Amana emblem
lg = crop(2, 862, 14, 910, 66); lg = lg.resize((lg.width * 3, lg.height * 3), Image.LANCZOS)
px = lg.load()
for x in range(lg.width):
    for y in range(lg.height):
        r, g, b = px[x, y]
        if min(r, g, b) > 225 and max(r, g, b) - min(r, g, b) < 16: px[x, y] = (255, 255, 255)
lg.save(f"{out}/logo-amana.png")

def dotted(path, w=1920, h=1358, step=28):
    im = Image.new("RGB", (w, h), (255, 255, 255)); d = ImageDraw.Draw(im)
    for r, y in enumerate(range(0, h + step, step)):
        for x in range(-step, w + step, step):
            cx = x + (step // 2 if r % 2 else 0); d.ellipse((cx - 4, y - 4, cx + 4, y + 4), fill=(243, 246, 244))
    im.save(path)
dotted(f"{out}/bg-dots.png")
Image.new("RGB", (64, 64), (0, 106, 70)).save(f"{out}/bg-green.png")
