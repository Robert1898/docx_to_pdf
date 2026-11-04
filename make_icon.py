"""Genereaza icon.ico / icon.png pentru Document Converter.

Foaie alba cu colt indoit pe un patrat rotunjit indigo-violet, plus o insigna portocalie cu
doua sageti circulare (simbolul „conversie”). Pentru 16 si 24 px se deseneaza o varianta
simplificata (fara randuri de text, insigna mai mare), ca sa ramana lizibila in taskbar.
"""

import math

from PIL import Image, ImageDraw, ImageFilter

S = 1024  # desenam mare si micsoram pentru margini netede
INDIGO, VIOLET = (59, 91, 219), (124, 58, 237)
AMBER, ORANGE = (251, 176, 36), (234, 88, 12)
SMALL_SIZES = (16, 24)


def diagonal_gradient(size, start, end):
    w, h = size
    grad = Image.new("RGBA", size)
    px = grad.load()
    for y in range(h):
        for x in range(w):
            t = (x + y) / (w + h - 2)
            px[x, y] = (*(int(a + (b - a) * t) for a, b in zip(start, end, strict=True)), 255)
    return grad


def shadow(canvas, mask_img, offset, blur, alpha):
    sh = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    a = mask_img.getchannel("A").point(lambda v: v * alpha // 255)
    sh.paste((15, 15, 40, 255), offset, a)
    canvas.alpha_composite(sh.filter(ImageFilter.GaussianBlur(blur)))


def background(canvas):
    margin, radius = 40, 230
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((margin, margin, S - margin, S - margin),
                                           radius=radius, fill=255)
    tile = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    tile.paste(diagonal_gradient((S, S), INDIGO, VIOLET), (0, 0), mask)
    # lumina discreta in coltul de sus
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse((-300, -420, 760, 420), fill=(255, 255, 255, 40))
    glow.putalpha(Image.composite(glow.getchannel("A"), Image.new("L", (S, S), 0), mask))
    tile.alpha_composite(glow.filter(ImageFilter.GaussianBlur(60)))
    shadow(canvas, tile, (0, 18), 26, 90)
    canvas.alpha_composite(tile)


def page(canvas, simple):
    x0, y0, x1, y1 = (250, 170, 690, 760) if not simple else (170, 130, 600, 720)
    w = x1 - x0
    fold, radius = int(w * 0.30), 44
    sheet = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    mask = Image.new("L", (S, S), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((x0, y0, x1, y1), radius=radius, fill=255)
    md.polygon([(x1 - fold, y0 - 2), (x1 + 2, y0 - 2), (x1 + 2, y0 + fold)], fill=0)
    sheet.paste((255, 255, 255, 255), (0, 0), mask)
    d = ImageDraw.Draw(sheet)
    # coltul indoit, in ton cu fundalul
    d.polygon([(x1 - fold, y0), (x1 - fold, y0 + fold - radius // 2),
               (x1 - fold + radius // 2, y0 + fold), (x1, y0 + fold)], fill=(199, 210, 254, 255))
    if not simple:
        lw = 34
        for i, frac in enumerate((0.62, 0.74, 0.50, 0.66)):
            y = y0 + 250 + i * 72
            d.rounded_rectangle((x0 + 60, y, x0 + 60 + int((w - 120) * frac), y + lw),
                                radius=lw // 2, fill=(165, 180, 252, 255))
        d.rounded_rectangle((x0 + 60, y0 + 120, x0 + 60 + int((w - 120) * 0.45), y0 + 170),
                            radius=25, fill=(99, 102, 241, 255))
    shadow(canvas, sheet, (0, 16), 22, 110)
    canvas.alpha_composite(sheet)


def arrow_arc(d, cx, cy, r, start, end, width, head):
    """Arc in sensul acelor de ceasornic (unghiuri PIL) cu varf de sageata la capat.

    In PIL grosimea unui arc creste spre interior, deci desenam conturul la r + width/2
    ca linia mediana sa fie exact pe raza r, unde e si varful sagetii."""
    outer = r + width / 2
    d.arc((cx - outer, cy - outer, cx + outer, cy + outer), start, end, fill="white",
          width=width)
    a = math.radians(end)
    px, py = cx + r * math.cos(a), cy + r * math.sin(a)
    tx, ty = -math.sin(a), math.cos(a)  # tangenta (sensul arcului)
    nx, ny = math.cos(a), math.sin(a)  # radial
    d.polygon([
        (px + tx * head * 1.1, py + ty * head * 1.1),
        (px + nx * head, py + ny * head),
        (px - nx * head, py - ny * head),
    ], fill="white")


def badge(canvas, simple):
    cx, cy, r = (720, 730, 230) if not simple else (665, 665, 315)
    disc = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).ellipse((cx - r, cy - r, cx + r, cy + r), fill=255)
    disc.paste(diagonal_gradient((S, S), AMBER, ORANGE), (0, 0), mask)
    ring = ImageDraw.Draw(disc)
    ring.ellipse((cx - r, cy - r, cx + r, cy + r), outline=(255, 255, 255, 255),
                 width=26 if not simple else 34)
    ar, width, head = (int(r * 0.50), int(r * 0.15), int(r * 0.20))
    arrow_arc(ring, cx, cy, ar, 195, 315, width, head)
    arrow_arc(ring, cx, cy, ar, 15, 135, width, head)
    shadow(canvas, disc, (0, 14), 20, 120)
    canvas.alpha_composite(disc)


def render(simple=False):
    canvas = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    background(canvas)
    page(canvas, simple)
    badge(canvas, simple)
    return canvas


def main():
    full, small = render(), render(simple=True)
    full.save("icon.png")
    frames = [
        (small if s in SMALL_SIZES else full).resize((s, s), Image.LANCZOS)
        for s in (16, 24, 32, 48, 64, 128, 256)
    ]
    frames[-1].save("icon.ico", sizes=[f.size for f in frames], append_images=frames[:-1])


if __name__ == "__main__":
    main()
