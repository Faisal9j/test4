"""Generates brand assets (dotted background, gradients, cover photo) from the Ministry report PDF render."""
import sys, os, random
from PIL import Image, ImageDraw, ImageFilter
out = sys.argv[1] if len(sys.argv) > 1 else "assets"
cover_src = sys.argv[2] if len(sys.argv) > 2 else None
os.makedirs(out, exist_ok=True)

def slim(img, path, width):
    """downscale + palette-quantize an RGBA photo so the deck (which embeds it per slide) stays small"""
    img = img.resize((width, int(img.height * width / img.width)), Image.LANCZOS)
    img.quantize(256, method=Image.Quantize.FASTOCTREE, dither=Image.Dither.FLOYDSTEINBERG).save(path, optimize=True)

def hexc(h): return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def dotted(path, w=1920, h=1080, step=28):
    im = Image.new("RGB", (w, h), (255, 255, 255)); d = ImageDraw.Draw(im)
    for r, y in enumerate(range(0, h + step, step)):
        for x in range(-step, w + step, step):
            cx = x + (step // 2 if r % 2 else 0)
            d.ellipse((cx - 5, y - 5, cx + 5, y + 5), fill=(241, 245, 243))
    im.save(path)

def gradient(path, c1, c2, w=1920, h=1080, diag=True):
    a, b = hexc(c1), hexc(c2)
    im = Image.new("RGB", (w, h)); px = im.load()
    for y in range(h):
        for x in range(w):
            t = (x / w * 0.35 + y / h * 0.65) if diag else y / h
            px[x, y] = tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    im.save(path)

def rounded_gradient(path, c1, c2, w, h, r):
    s = 2; W, H = w * s, h * s
    g = Image.new("RGB", (W, H)); a, b = hexc(c1), hexc(c2); px = g.load()
    for y in range(H):
        for x in range(W):
            t = x / W * 0.5 + y / H * 0.5
            px[x, y] = tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))
    m = Image.new("L", (W, H), 0); ImageDraw.Draw(m).rounded_rectangle((0, 0, W - 1, H - 1), r * s, fill=255)
    g.putalpha(m); g.resize((w, h), Image.LANCZOS).save(path)

dotted(f"{out}/bg-dots.png")
gradient(f"{out}/bg-dark.png", "006A46", "005A3C", diag=False)
rounded_gradient(f"{out}/card-green.png", "0A7A54", "004D33", 800, 400, 60)
rounded_gradient(f"{out}/panel-green.png", "0A7A54", "004D33", 1600, 1000, 70)
rounded_gradient(f"{out}/panel-teal.png", "1C8A64", "006A46", 1100, 1000, 90)
if cover_src:
    im = Image.open(cover_src).convert("RGB"); w, h = im.size; sx, sy = w / 1008, h / 2000
    c = im.crop((int(46 * sx), int(86 * sy), int(962 * sx), int(1215 * sy))).convert("RGBA")
    W, H = c.size; r = int(W * 0.07); m = Image.new("L", (W, H), 255); d = ImageDraw.Draw(m)
    for box in [(0, 0, 2 * r, 2 * r), (W - 2 * r, H - 2 * r, W, H), (0, H - 2 * r, 2 * r, H)]:
        pass
    m = Image.new("L", (W * 2, H * 2), 0); ImageDraw.Draw(m).rounded_rectangle((0, 0, W * 2 - 1, H * 2 - 1), r * 2, fill=255)
    c.putalpha(m.resize((W, H), Image.LANCZOS)); slim(c, f"{out}/cover-photo.png", 1200)


# ---- portrait assets + notched photo frames (photos cropped from the Ministry report pages)
def dotted_p(path, w=1500, h=2980, step=28):
    im = Image.new("RGB", (w, h), (255, 255, 255)); d = ImageDraw.Draw(im)
    for r, y in enumerate(range(0, h + step, step)):
        for x in range(-step, w + step, step):
            cx = x + (step // 2 if r % 2 else 0)
            d.ellipse((cx - 5, y - 5, cx + 5, y + 5), fill=(241, 245, 243))
    im.save(path)
dotted_p(f"{out}/bg-dots-p.png")
gradient(f"{out}/bg-dark-p.png", "006A46", "005A3C", w=1500, h=2980, diag=False)

def notch_frame(img, path, w=1340, h=520, r=60, nw=640, nh=200):
    img = img.convert("RGB"); iw, ih = img.size
    t = w / h
    if iw / ih > t: nw_ = int(ih * t); img = img.crop(((iw - nw_) // 2, 0, (iw - nw_) // 2 + nw_, ih))
    else: nh_ = int(iw / t); img = img.crop((0, (ih - nh_) // 2, iw, (ih - nh_) // 2 + nh_))
    img = img.resize((w, h), Image.LANCZOS).convert("RGBA")
    s = 2; m = Image.new("L", (w * s, h * s), 0); d = ImageDraw.Draw(m)
    d.rounded_rectangle((0, 0, w * s - 1, h * s - 1), r * s, fill=255)
    d.rounded_rectangle((w * s - nw * s, -r * s, w * s + r * s, nh * s), r * s, fill=0)   # notch top-right
    img.putalpha(m.resize((w, h), Image.LANCZOS)); slim(img, path, 1000)

if len(sys.argv) > 3:
    base = sys.argv[3]
    def pg(n): return Image.open(f"{base}-{n:02d}.png").convert("RGB")
    def box(im, x0, y0, x1, y1):
        W, H = im.size; sx, sy = W / 1008, H / 2000
        return im.crop((int(x0 * sx), int(y0 * sy), int(x1 * sx), int(y1 * sy)))
    notch_frame(box(pg(2), 90, 470, 917, 815), f"{out}/ph0.png")
    for i, n in enumerate((14, 15, 16, 17), 1):
        notch_frame(box(pg(n), 90, 332, 510, 585), f"{out}/ph{i}.png")
