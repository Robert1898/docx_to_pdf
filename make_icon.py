"""Genereaza icon.ico pentru executabil (document Word -> document PDF)."""

from PIL import Image, ImageDraw, ImageFilter, ImageFont

S = 1024  # desenam mare si micsoram pentru margini netede
FONT = r"C:\Windows\Fonts\segoeuib.ttf"


def font(size):
    return ImageFont.truetype(FONT, size)


def vertical_gradient(size, top, bottom):
    w, h = size
    grad = Image.new("RGBA", (1, h))
    for y in range(h):
        t = y / (h - 1)
        grad.putpixel((0, y), tuple(int(a + (b - a) * t) for a, b in zip(top, bottom)) + (255,))
    return grad.resize((w, h))


def document(size, top, bottom, label, label_size, label_y, lines=True, fold=0.26):
    """Foaie cu colt indoit, umplere in gradient si o eticheta alba."""
    w, h = size
    f = int(w * fold)
    r = int(w * 0.08)
    mask = Image.new("L", size, 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((0, 0, w - 1, h - 1), radius=r, fill=255)
    md.polygon([(w - f, 0), (w, 0), (w, f)], fill=0)

    doc = Image.new("RGBA", size, (0, 0, 0, 0))
    doc.paste(vertical_gradient(size, top, bottom), (0, 0), mask)
    d = ImageDraw.Draw(doc)

    # coltul indoit
    d.polygon([(w - f, 0), (w - f, f - r // 2), (w - f + r // 2, f), (w, f)],
              fill=tuple(min(255, c + 70) for c in top) + (255,))

    # randuri de "text", desenate pe un strat separat ca sa se amestece corect
    if lines:
        overlay = Image.new("RGBA", size, (0, 0, 0, 0))
        od = ImageDraw.Draw(overlay)
        lw = int(w * 0.06)
        for i, frac in enumerate((0.60, 0.72, 0.48)):
            y = int(h * 0.24) + i * int(h * 0.10)
            od.rounded_rectangle((int(w * 0.14), y, int(w * (0.14 + frac)), y + lw),
                                 radius=lw // 2, fill=(255, 255, 255, 120))
        doc.alpha_composite(overlay)

    fnt = font(label_size)
    bbox = d.textbbox((0, 0), label, font=fnt)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    d.text(((w - tw) / 2 - bbox[0], h * label_y - th / 2 - bbox[1]), label, font=fnt, fill="white")
    return doc


def paste_with_shadow(canvas, img, pos, offset=(12, 20), blur=24, alpha=120):
    # umbra se deseneaza pe tot canvasul, ca blur-ul sa nu fie taiat la margini
    sh = Image.new("RGBA", canvas.size, (0, 0, 0, 0))
    a = img.getchannel("A").point(lambda v: v * alpha // 255)
    sh.paste((0, 0, 0, 255), (pos[0] + offset[0], pos[1] + offset[1]), a)
    canvas.alpha_composite(sh.filter(ImageFilter.GaussianBlur(blur)))
    canvas.alpha_composite(img, pos)


def main():
    canvas = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    word = document((500, 640), (66, 133, 244), (24, 78, 190), "W", 280, 0.42, lines=False)
    pdf = document((560, 700), (244, 81, 72), (190, 30, 45), "PDF", 180, 0.76)

    paste_with_shadow(canvas, word, (50, 30))
    paste_with_shadow(canvas, pdf, (420, 290))

    # sageata intr-un cerc, de la Word spre PDF
    cx, cy, cr = 250, 800, 160
    circle = Image.new("RGBA", (2 * cr, 2 * cr), (0, 0, 0, 0))
    cd = ImageDraw.Draw(circle)
    cd.ellipse((0, 0, 2 * cr - 1, 2 * cr - 1), fill="white")
    cd.ellipse((20, 20, 2 * cr - 21, 2 * cr - 21), fill=(46, 196, 120, 255))
    shaft, head = 40, 78
    cd.rounded_rectangle((cr - 95, cr - shaft // 2, cr + 25, cr + shaft // 2), radius=12, fill="white")
    cd.polygon([(cr + 5, cr - head), (cr + 105, cr), (cr + 5, cr + head)], fill="white")
    paste_with_shadow(canvas, circle, (cx - cr, cy - cr), offset=(8, 14), blur=18)

    canvas.save("icon.png")
    canvas.save("icon.ico", sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])


if __name__ == "__main__":
    main()
