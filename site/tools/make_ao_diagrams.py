#!/usr/bin/env python3
"""Draw the Angola diagrams used on the /ao/ and /ao/pt/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_ao_diagrams.py

Writes two schematics to site/images/diagrams/, with no words but distances, so one image serves both
languages and the page caption names the zones and the laws:

- ao-strips.png: a plan view of a pipeline or line axis drawn to scale, with the 30 m strip that
  Lei n.º 9/04 art. 27(7)(g) sets on each side of electricity, water, telecom, oil and gas
  installations and conductors (red), and two wider bands a client may add to a survey, 60 m (amber)
  and 100 m (teal, dashed). The structures are invented and coloured by band.
- ao-mining-zones.png: a mining area drawn to scale with the two zones of the Código Mineiro
  (Lei n.º 31/11): the zona restrita, up to a 1 000 m radius around the deposits and plants
  (art. 200), and the zona de protecção, up to 5 km from the outer limits of the protected deposits
  (art. 202). Invented excavations, cleared patches, tracks and structures show the kinds of change
  a survey maps. Both zones are drawn at their maximum width; the real ones are set for each area.

Deterministic: the same code always draws the same pixels.
"""
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SITE = Path(__file__).resolve().parents[1]
OUT = SITE / "images" / "diagrams"
FONT = SITE / "fonts" / "inter-latin-wght-normal.woff2"
SS = 2
W, H = 1600, 1000

BG, GRID, TEXT, DIM, LINE = "#ffffff", "#eef1f4", "#172030", "#4b5563", "#c9d0d9"
ORANGE, RED, AMBER, TEAL, BROWN, SLATE = "#e8672f", "#dc2626", "#f59e0b", "#0f766e", "#7c4a1e", "#64748b"


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

    def poly(self, pts, fill=None, outline=None, width=2):
        layer = Image.new("RGBA", self.im.size, (0, 0, 0, 0))
        d = ImageDraw.Draw(layer)
        d.polygon([(x * SS, y * SS) for x, y in pts], fill=fill)
        self.im.alpha_composite(layer)
        if outline:
            self.line(pts + [pts[0]], outline, width)

    def line(self, pts, color, width=2):
        self.d.line([(x * SS, y * SS) for x, y in pts], fill=color, width=width * SS, joint="curve")

    def dashed(self, pts, color, width=3, dash=22, gap=14):
        """A dashed polyline whose dash pattern runs on across vertices (so curves stay dashed)."""
        on, left = True, dash
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            seg = math.hypot(x1 - x0, y1 - y0)
            t = 0.0
            while t < seg:
                step = min(left, seg - t)
                if on:
                    self.line([(x0 + (x1 - x0) * t / seg, y0 + (y1 - y0) * t / seg),
                               (x0 + (x1 - x0) * (t + step) / seg, y0 + (y1 - y0) * (t + step) / seg)], color, width)
                t += step
                left -= step
                if left <= 1e-9:
                    on = not on
                    left = dash if on else gap

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

    def arrow(self, a, b, color, width=4, head=18):
        self.line([a, b], color, width)
        for (px, py), (qx, qy) in ((a, b), (b, a)):
            ang = math.atan2(qy - py, qx - px)
            left = (px + head * math.cos(ang + 0.42), py + head * math.sin(ang + 0.42))
            right = (px + head * math.cos(ang - 0.42), py + head * math.sin(ang - 0.42))
            self.d.polygon([(px * SS, py * SS), (left[0] * SS, left[1] * SS), (right[0] * SS, right[1] * SS)],
                           fill=color)

    def frame(self):
        mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
        self.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
        for x in range(mx0 + 40, mx1 - 20, 64):
            self.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
        for y in range(my0 + 40, my1 - 20, 64):
            self.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)
        return mx0, my0, mx1, my1

    def save(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.im.resize((W, H), Image.LANCZOS).convert("RGB").save(path, optimize=True)
        print("wrote", path.relative_to(SITE.parent))


def strips():
    c = Canvas()
    mx0, my0, mx1, my1 = c.frame()
    k = 4.0                                    # pixels per metre: 100 m each side fits the panel
    axis = (my0 + my1) / 2
    bx0, bx1 = 540, mx1 - 36
    y = lambda m: axis - m * k

    for side in (1, -1):
        c.fill([bx0, min(y(100 * side), y(60 * side)), bx1, max(y(100 * side), y(60 * side))], rgba(TEAL, 26))
        c.fill([bx0, min(y(60 * side), y(30 * side)), bx1, max(y(60 * side), y(30 * side))], rgba(AMBER, 70))
        c.fill([bx0, min(y(30 * side), axis), bx1, max(y(30 * side), axis)], rgba(RED, 60))
        c.line([(bx0, y(30 * side)), (bx1, y(30 * side))], rgba(RED, 200), 3)
        c.line([(bx0, y(60 * side)), (bx1, y(60 * side))], rgba("#b45309", 170), 3)
        c.dashed([(bx0, y(100 * side)), (bx1, y(100 * side))], rgba(TEAL, 230), 4)

    # invented structures (x, metres from the axis), coloured by band
    plan = [(590, 21), (680, -12), (760, 26), (850, -24), (990, 9), (1250, -19), (1400, 17),
            (630, 44), (720, -52), (905, 38), (1060, -41), (1170, 55), (1310, -47), (1465, 35),
            (610, 83), (800, -76), (965, 92), (1115, -90), (1275, 71), (1425, -66), (1505, 88)]
    for x, m in plan:
        a = abs(m)
        c.square((x, y(m)), RED if a <= 30 else AMBER if a <= 60 else TEAL)

    c.line([(bx0 - 10, axis), (bx1, axis)], ORANGE, 10)

    for gx, m, color, label in ((390, 30, RED, "30 m"), (250, 60, "#b45309", "60 m"), (110, 100, TEAL, "100 m")):
        c.arrow_v(gx, y(m), axis, color, 5, 18)
        c.line([(gx - 18, axis), (bx0 - 10, axis)], rgba(DIM, 120), 2)
        c.line([(gx - 18, y(m)), (bx0, y(m))], rgba(color, 150), 2)
        c.text((gx + 12, y(m) - 12), label, 54, 800, color, anchor="lb")
    c.save(OUT / "ao-strips.png")


def offset_polygon(pts, d, steps=96):
    """Outward offset of a convex polygon by d pixels: the convex hull of a circle of radius d swept
    round every vertex (the Minkowski sum), which gives the rounded corners of a real buffer."""
    cloud = [(x + d * math.cos(2 * math.pi * i / steps), y + d * math.sin(2 * math.pi * i / steps))
             for x, y in pts for i in range(steps)]
    cloud = sorted(set(cloud))

    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower, upper = [], []
    for p in cloud:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    for p in reversed(cloud):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]


def mining_zones():
    c = Canvas()
    mx0, my0, mx1, my1 = c.frame()
    km = 66.0                                  # pixels per kilometre
    cx, cy = 800, 505
    deposit = [(cx - 110, cy - 70), (cx - 20, cy - 118), (cx + 105, cy - 84), (cx + 150, cy + 10),
               (cx + 80, cy + 96), (cx - 60, cy + 104), (cx - 140, cy + 30)]
    protection = offset_polygon(deposit, 5 * km)
    restricted = offset_polygon(deposit, 1 * km)

    c.poly(protection, fill=rgba(AMBER, 46))
    c.dashed(protection + [protection[0]], rgba("#b45309", 220), 4)
    c.poly(restricted, fill=rgba(RED, 58))
    c.line(restricted + [restricted[0]], rgba(RED, 210), 3)
    c.poly(deposit, fill=rgba(SLATE, 150), outline=rgba("#334155", 230), width=3)
    c.rect([cx + 20, cy - 40, cx + 70, cy + 10], fill="#334155", outline="#ffffff", width=3, radius=4)

    # invented change: new tracks (dashed brown), excavations (brown blobs), cleared ground (hatched)
    c.dashed([(310, 820), (420, 720), (560, 660), (650, 610)], rgba(BROWN, 230), 5, 18, 12)
    c.dashed([(1330, 180), (1210, 290), (1090, 350), (980, 430)], rgba(BROWN, 230), 5, 18, 12)
    for (x, y, r) in ((600, 640, 16), (628, 668, 11), (1000, 690, 14), (1030, 712, 10), (455, 320, 13),
                      (1180, 560, 15), (1205, 585, 10)):
        c.d.ellipse([(x - r) * SS, (y - r * 0.7) * SS, (x + r) * SS, (y + r * 0.7) * SS],
                    fill=rgba(BROWN, 200), outline=rgba("#ffffff", 255), width=2 * SS)
    for (x0, y0, x1, y1) in ((1060, 740, 1160, 810), (470, 395, 550, 455)):
        c.fill([x0, y0, x1, y1], rgba(BROWN, 40))
        for t in range(0, int((x1 - x0) + (y1 - y0)), 16):
            a = (x0 + t, y0) if t <= x1 - x0 else (x1, y0 + t - (x1 - x0))
            b = (x0, y0 + t) if t <= y1 - y0 else (x0 + t - (y1 - y0), y1)
            c.line([a, b], rgba(BROWN, 120), 2)
        c.rect([x0, y0, x1, y1], outline=rgba(BROWN, 180), width=2)

    # invented structures, coloured by zone
    for x, y in ((700, 600), (905, 385), (655, 455), (940, 610)):
        c.square((x, y), RED)
    for x, y in ((600, 300), (1150, 470), (560, 760), (1080, 250), (360, 560), (1250, 700), (820, 850),
                 (780, 160), (1300, 420)):
        c.square((x, y), AMBER)
    for x, y in ((120, 150), (170, 190), (1480, 880), (1440, 845), (1500, 150), (140, 880)):
        c.square((x, y), TEAL)

    # distances, measured from the deposit's outer limit
    ex, ey = cx + 150, cy + 10
    c.arrow((ex, ey), (ex + 1 * km, ey), RED, 5)
    c.text((ex + 1 * km + 14, ey - 34), "1 km", 50, 800, RED, anchor="la")
    c.arrow((cx - 140, cy + 30), (cx - 140 - 5 * km, cy + 30), "#b45309", 5)
    c.text((cx - 140 - 5 * km + 24, cy + 18), "5 km", 50, 800, "#b45309", anchor="lb")
    c.save(OUT / "ao-mining-zones.png")


if __name__ == "__main__":
    strips()
    mining_zones()
