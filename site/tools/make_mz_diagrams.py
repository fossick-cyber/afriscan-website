#!/usr/bin/env python3
"""Draw the Mozambique legal-strip diagram used on the /mz/ and /mz/pt/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_mz_diagrams.py

Writes site/images/diagrams/mz-strips.png: a plan view of a pipeline axis drawn to scale, with the
50 m partial protection zone (Lei n.º 19/97 art. 8(g); Lei n.º 8/2026 art. 75(3)), the 100 m outer
band AfriScan reports by default, and a 200 m safety zone of the kind a decree can set for a
corridor. The only words in the image are the three widths; the page caption names the zones and the
laws, so the image needs no translation and stays readable on a phone. The structures
are invented and coloured by the same bands the register uses (red within 50 m, amber 50–100 m, teal
beyond). Deterministic: the same code always draws the same pixels.
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SITE = Path(__file__).resolve().parents[1]
OUT = SITE / "images" / "diagrams"
FONT = SITE / "fonts" / "inter-latin-wght-normal.woff2"
SS = 2
W, H = 1600, 1000

BG, GRID, TEXT, DIM, LINE = "#ffffff", "#eef1f4", "#172030", "#4b5563", "#c9d0d9"
ORANGE, RED, AMBER, TEAL = "#e8672f", "#dc2626", "#f59e0b", "#0f766e"


def font(size, weight=600):
    f = ImageFont.truetype(str(FONT), size * SS)
    try:
        f.set_variation_by_axes([weight])
    except Exception:
        pass
    return f


def rgba(hex_, a=255):
    hex_ = hex_.lstrip("#")
    return tuple(int(hex_[i:i + 2], 16) for i in (0, 2, 4)) + (a,)


class Canvas:
    def __init__(self):
        self.im = Image.new("RGBA", (W * SS, H * SS), rgba(BG))
        self.d = ImageDraw.Draw(self.im, "RGBA")

    def rect(self, box, fill=None, outline=None, width=1, radius=0):
        self.d.rounded_rectangle([v * SS for v in box], radius=radius * SS, fill=fill, outline=outline,
                                 width=width * SS)

    def fill(self, box, color):
        layer = Image.new("RGBA", self.im.size, (0, 0, 0, 0))
        ImageDraw.Draw(layer).rectangle([v * SS for v in box], fill=color)
        self.im.alpha_composite(layer)

    def line(self, pts, color, width=2):
        self.d.line([(x * SS, y * SS) for x, y in pts], fill=color, width=width * SS)

    def dashed_h(self, x0, x1, y, color, width=3, dash=22, gap=14):
        x = x0
        while x < x1:
            self.line([(x, y), (min(x + dash, x1), y)], color, width)
            x += dash + gap

    def text(self, xy, s, size=40, weight=700, color=TEXT, anchor="la"):
        self.d.text((xy[0] * SS, xy[1] * SS), s, font=font(size, weight), fill=color, anchor=anchor)

    def square(self, c, color, size=26):
        x, y = c
        h = size / 2
        self.rect([x - h, y - h, x + h, y + h], fill=color, outline="#ffffff", width=3, radius=3)

    def arrow_v(self, x, y0, y1, color, width=4, head=16):
        self.line([(x, y0), (x, y1)], color, width)
        for y, s in ((y0, 1), (y1, -1)):
            self.d.polygon([(x * SS, y * SS), ((x - head / 2) * SS, (y + s * head) * SS),
                            ((x + head / 2) * SS, (y + s * head) * SS)], fill=color)

    def save(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.im.resize((W, H), Image.LANCZOS).convert("RGB").save(path, optimize=True)
        print("wrote", path.relative_to(SITE.parent))


def strips():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    k = 2.05                                   # pixels per metre: 200 m each side fits the panel
    axis = (my0 + my1) / 2
    bx0, bx1 = 540, mx1 - 36                   # bands run from the dimension gutter to the right edge
    y = lambda m: axis - m * k                 # metres from the axis (negative = the other side)

    for side in (1, -1):
        c.fill([bx0, min(y(200 * side), y(100 * side)), bx1, max(y(200 * side), y(100 * side))], rgba(TEAL, 26))
        c.fill([bx0, min(y(100 * side), y(50 * side)), bx1, max(y(100 * side), y(50 * side))], rgba(AMBER, 70))
        c.fill([bx0, min(y(50 * side), axis), bx1, max(y(50 * side), axis)], rgba(RED, 60))
        c.line([(bx0, y(50 * side)), (bx1, y(50 * side))], rgba(RED, 190), 3)
        c.line([(bx0, y(100 * side)), (bx1, y(100 * side))], rgba("#b45309", 170), 3)
        c.dashed_h(bx0, bx1, y(200 * side), rgba(TEAL, 230), 4)

    # invented structures at set offsets (metres from the axis), coloured by band
    plan = [(600, 32), (690, -18), (770, 44), (840, -41), (985, 12), (1240, -30), (1390, 27),
            (640, 72), (730, -88), (900, 64), (1060, -66), (1170, 91), (1310, -79), (1460, 58),
            (620, 150), (800, -140), (960, 175), (1110, -182), (1270, 128), (1420, -120), (1500, 186)]
    for x, m in plan:
        a = abs(m)
        c.square((x, y(m)), RED if a <= 50 else AMBER if a <= 100 else TEAL)

    c.line([(bx0 - 10, axis), (bx1, axis)], ORANGE, 10)

    # dimension arrows in the left gutter, measured from the axis
    for gx, m, color, label in ((390, 50, RED, "50 m"), (250, 100, "#b45309", "100 m"), (110, 200, TEAL, "200 m")):
        c.arrow_v(gx, y(m), axis, color, 5, 18)
        c.line([(gx - 18, axis), (bx0 - 10, axis)], rgba(DIM, 120), 2)
        c.line([(gx - 18, y(m)), (bx0, y(m))], rgba(color, 150), 2)
        c.text((gx + 12, y(m) - 12), label, 54, 800, color, anchor="lb")
    c.save(OUT / "mz-strips.png")


if __name__ == "__main__":
    strips()
