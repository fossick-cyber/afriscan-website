#!/usr/bin/env python3
"""Draw the South Africa schematics used on the /za/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_za_diagrams.py

Writes to site/images/diagrams/:
  za-road-restriction.png  SANRAL Act 7 of 1998 s48: the building restriction area is land outside
                           urban areas within 60 m of a national road's boundary, or within 500 m of a
                           point of intersection. Not to scale: the 60 m strip is drawn wider than the
                           500 m radius would make it, so it reads on a phone; the reserve width is
                           illustrative.
  za-servitude-change.png  The same stretch of a power-line servitude on an earlier and a later survey:
                           structures that are new on the later survey carry a ring, one removed
                           structure is drawn as a dashed outline, and cleared ground and a new track
                           appear. It illustrates how change between dated surveys is reported.
The only words in the images are dimension labels and the two panel titles; the page captions carry
the legend, so the images stay readable on a phone. Geometry and structures are invented.
Deterministic: the same code always draws the same pixels.
"""
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SITE = Path(__file__).resolve().parents[1]
OUT = SITE / "images" / "diagrams"
FONT = SITE / "fonts" / "inter-latin-wght-normal.woff2"
SS = 2

BG, GRID, TEXT, DIM, LINE = "#ffffff", "#eef1f4", "#172030", "#4b5563", "#c9d0d9"
ORANGE, RED, AMBER, TEAL, GREY, BROWN = "#e8672f", "#dc2626", "#f59e0b", "#0f766e", "#9aa3ae", "#8a5a2b"


def font(size, weight=700):
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
    def __init__(self, w, h):
        self.w, self.h = w, h
        self.im = Image.new("RGBA", (w * SS, h * SS), rgba(BG))
        self.d = ImageDraw.Draw(self.im, "RGBA")

    def _layer(self, draw_fn):
        layer = Image.new("RGBA", self.im.size, (0, 0, 0, 0))
        draw_fn(ImageDraw.Draw(layer))
        self.im.alpha_composite(layer)

    def rect(self, box, fill=None, outline=None, width=1, radius=0):
        self.d.rounded_rectangle([v * SS for v in box], radius=radius * SS, fill=fill, outline=outline,
                                 width=width * SS)

    def fill_rect(self, box, color):
        self._layer(lambda d: d.rectangle([v * SS for v in box], fill=color))

    def fill_poly(self, pts, color):
        self._layer(lambda d: d.polygon([(x * SS, y * SS) for x, y in pts], fill=color))

    def fill_circle(self, c, r, color):
        x, y = c
        self._layer(lambda d: d.ellipse([(x - r) * SS, (y - r) * SS, (x + r) * SS, (y + r) * SS], fill=color))

    def line(self, pts, color, width=2):
        self.d.line([(x * SS, y * SS) for x, y in pts], fill=color, width=width * SS, joint="curve")

    def dashed(self, pts, color, width=3, dash=18, gap=12):
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            L = math.hypot(x1 - x0, y1 - y0)
            pos = 0.0
            while pos < L:
                e = min(pos + dash, L)
                self.line([(x0 + (x1 - x0) * pos / L, y0 + (y1 - y0) * pos / L),
                           (x0 + (x1 - x0) * e / L, y0 + (y1 - y0) * e / L)], color, width)
                pos += dash + gap

    def dashed_circle(self, c, r, color, width=4, n=72):
        x, y = c
        for i in range(0, n, 2):
            a0, a1 = 2 * math.pi * i / n, 2 * math.pi * (i + 1) / n
            self.line([(x + r * math.cos(a0), y + r * math.sin(a0)), (x + r * math.cos(a1), y + r * math.sin(a1))],
                      color, width)

    def text(self, xy, s, size=40, weight=700, color=TEXT, anchor="la"):
        self.d.text((xy[0] * SS, xy[1] * SS), s, font=font(size, weight), fill=color, anchor=anchor)

    def label(self, xy, s, size=38, color=TEXT, anchor="mm"):
        """Text on a white pill so it stays legible over tinted bands."""
        f = font(size, 800)
        l, t, r, b = self.d.textbbox((xy[0] * SS, xy[1] * SS), s, font=f, anchor=anchor)
        pad = 10 * SS
        self.d.rounded_rectangle([l - pad, t - pad * 0.7, r + pad, b + pad * 0.7], radius=8 * SS,
                                 fill=rgba("#ffffff", 235))
        self.d.text((xy[0] * SS, xy[1] * SS), s, font=f, fill=color, anchor=anchor)

    def square(self, c, color, size=24, ring=None, ghost=False):
        x, y = c
        h = size / 2
        if ghost:
            self.dashed([(x - h, y - h), (x + h, y - h), (x + h, y + h), (x - h, y + h), (x - h, y - h)],
                        rgba(DIM, 220), 3, 7, 5)
            return
        if ring:
            self.rect([x - h - 8, y - h - 8, x + h + 8, y + h + 8], outline=ring, width=4, radius=5)
        self.rect([x - h, y - h, x + h, y + h], fill=color, outline="#ffffff", width=3, radius=3)

    def grid(self, box, step=64):
        x0, y0, x1, y1 = box
        for x in range(x0, x1 + 1, step):
            self.line([(x, y0), (x, y1)], GRID, 1)
        for y in range(y0, y1 + 1, step):
            self.line([(x0, y), (x1, y)], GRID, 1)

    def save(self, name):
        OUT.mkdir(parents=True, exist_ok=True)
        path = OUT / name
        self.im.resize((self.w, self.h), Image.LANCZOS).convert("RGB").save(path, optimize=True)
        print("wrote", path.relative_to(SITE.parent))


# ---------------------------------------------------------------- SANRAL building restriction area
def road_restriction():
    W, H = 1600, 1000
    c = Canvas(W, H)
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    c.grid((mx0 + 20, my0 + 20, mx1 - 20, my1 - 20))

    k = 0.72                                   # px per metre for the 500 m radius
    axis = 470                                 # national road centreline
    half_reserve = 40                          # road reserve half-width in px (illustrative)
    top_b, bot_b = axis - half_reserve, axis + half_reserve
    strip = 60 * k * 1.6                       # 60 m drawn a little wider so it reads on a phone
    urban_x = 1230                             # urban area from here to the right edge
    ix = 560                                   # point of intersection (side road crosses here)
    r500 = 500 * k

    left, right = mx0 + 2, mx1 - 2
    # urban area (outside the building restriction area)
    c.fill_rect([urban_x, my0 + 2, right, my1 - 2], rgba(GREY, 70))
    # 500 m circle around the point of intersection, clipped to non-urban land
    circ = Image.new("RGBA", c.im.size, (0, 0, 0, 0))
    ImageDraw.Draw(circ).ellipse([(ix - r500) * SS, (axis - r500) * SS, (ix + r500) * SS, (axis + r500) * SS],
                                 fill=rgba(ORANGE, 38))
    mask = Image.new("L", c.im.size, 0)
    ImageDraw.Draw(mask).rectangle([left * SS, (my0 + 2) * SS, urban_x * SS, (my1 - 2) * SS], fill=255)
    clipped = Image.new("RGBA", c.im.size, (0, 0, 0, 0))
    clipped.paste(circ, (0, 0), mask)
    c.im.alpha_composite(clipped)
    # 60 m strips beyond each reserve boundary, outside the urban area
    c.fill_rect([left, top_b - strip, urban_x, top_b], rgba(ORANGE, 95))
    c.fill_rect([left, bot_b, urban_x, bot_b + strip], rgba(ORANGE, 95))
    c.line([(left, top_b - strip), (urban_x, top_b - strip)], rgba(ORANGE, 220), 3)
    c.line([(left, bot_b + strip), (urban_x, bot_b + strip)], rgba(ORANGE, 220), 3)
    c.dashed_circle((ix, axis), r500, rgba("#b4491c", 230), 4)

    # side road crossing the national road at the point of intersection
    c.fill_rect([ix - 14, my0 + 2, ix + 14, my1 - 2], rgba("#d5d9df", 255))
    c.line([(ix, my0 + 2), (ix, my1 - 2)], rgba("#ffffff", 255), 2)
    # national road reserve and carriageway
    c.fill_rect([left, top_b, right, bot_b], rgba("#d5d9df", 255))
    c.line([(left, top_b), (right, top_b)], rgba(DIM, 230), 3)
    c.line([(left, bot_b), (right, bot_b)], rgba(DIM, 230), 3)
    c.fill_rect([left, axis - 16, right, axis + 16], rgba("#5b6470", 255))
    c.dashed([(left, axis), (right, axis)], rgba("#ffffff", 255), 3, 26, 22)
    c.fill_rect([ix - 14, axis - 16, ix + 14, axis + 16], rgba("#5b6470", 255))
    c.fill_circle((ix, axis), 9, rgba("#ffffff", 255))

    # invented structures; red = inside the restriction area, teal = outside it
    def in_bra(x, y):
        if x >= urban_x:
            return False
        if top_b - strip <= y <= top_b or bot_b <= y <= bot_b + strip:
            return True
        return math.hypot(x - ix, y - axis) <= r500 and not (top_b < y < bot_b)

    pts = [(175, 395), (230, 380), (330, 560), (410, 395), (700, 372), (760, 560), (880, 388), (1000, 552),
           (1110, 384), (455, 250), (660, 190), (520, 720), (720, 690), (300, 640), (380, 170), (160, 250),
           (900, 250), (1020, 300), (960, 700), (1140, 640), (1080, 180), (240, 840), (820, 860), (1180, 860),
           (1290, 380), (1350, 560), (1420, 395), (1500, 250), (1310, 700), (1460, 640), (1380, 180), (1520, 820)]
    for x, y in pts:
        c.square((x, y), RED if in_bra(x, y) else TEAL)

    # dimension labels
    c.line([(80, top_b), (80, top_b - strip)], rgba("#b4491c", 255), 4)
    c.label((80, top_b - strip / 2), "60 m", 34, "#b4491c")
    c.line([(ix, axis), (ix + r500 * math.cos(-0.62), axis + r500 * math.sin(-0.62))], rgba("#b4491c", 255), 4)
    c.label((ix + r500 * 0.78 * math.cos(-0.62), axis + r500 * 0.78 * math.sin(-0.62)), "500 m", 34, "#b4491c")
    c.save("za-road-restriction.png")


# ---------------------------------------------------------------- servitude change between surveys
def servitude_change():
    W, H = 1800, 860
    c = Canvas(W, H)
    panels = [(24, 24, 888, H - 24, "Earlier survey"), (912, 24, W - 24, H - 24, "Later survey")]
    for pi, (x0, y0, x1, y1, title) in enumerate(panels):
        c.rect([x0, y0, x1, y1], fill=BG, outline=LINE, width=2, radius=18)
        c.grid((x0 + 20, y0 + 80, x1 - 20, y1 - 20), 60)
        c.text((x0 + 36, y0 + 26), title, 40, 800, TEXT)
        cy = (y0 + y1) / 2 + 30
        half = 120
        # servitude edges and band
        c.fill_rect([x0 + 2, cy - half, x1 - 2, cy + half], rgba(ORANGE, 30))
        c.dashed([(x0 + 2, cy - half), (x1 - 2, cy - half)], rgba(ORANGE, 230), 4)
        c.dashed([(x0 + 2, cy + half), (x1 - 2, cy + half)], rgba(ORANGE, 230), 4)
        # line and towers
        c.line([(x0 + 2, cy), (x1 - 2, cy)], rgba("#374151", 255), 4)
        for tx in (x0 + 120, x0 + 430, x0 + 740):
            c.d.polygon([((tx) * SS, (cy - 22) * SS), ((tx + 18) * SS, (cy + 18) * SS), ((tx - 18) * SS, (cy + 18) * SS)],
                        fill=rgba("#374151", 255))
        ox = x0
        base = [(ox + 250, cy - 60), (ox + 300, cy + 170), (ox + 560, cy + 70), (ox + 620, cy - 190),
                (ox + 690, cy + 200), (ox + 170, cy - 200), (ox + 800, cy - 150)]
        new = [(ox + 350, cy - 80), (ox + 395, cy - 30), (ox + 330, cy + 60), (ox + 505, cy - 95)]
        removed = (ox + 560, cy + 70)
        if pi == 1:
            # cleared ground and a new track that appear on the later survey
            c.fill_poly([(ox + 300, cy - 118), (ox + 540, cy - 118), (ox + 560, cy + 95), (ox + 290, cy + 95)],
                        rgba(BROWN, 60))
            c.dashed([(ox + 420, y0 + 82), (ox + 430, cy - 118), (ox + 440, cy + 95)], rgba(BROWN, 255), 5, 16, 10)
        for p in base:
            inside = abs(p[1] - cy) <= half
            if pi == 1 and p == removed:
                c.square(p, TEAL, ghost=True)
                continue
            c.square(p, RED if inside else TEAL)
        if pi == 1:
            for p in new:
                c.square(p, RED, ring=TEXT)
    c.save("za-servitude-change.png")


if __name__ == "__main__":
    road_restriction()
    servitude_change()
