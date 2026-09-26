#!/usr/bin/env python3
"""Draw the schematic diagrams used on the industry pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_diagrams.py

Writes site/images/diagrams/{corridor,area-ring}.png. There is no legend inside the image: the
page carries it as text in the figure caption, so it stays readable on a phone. They are
schematics, not real sites: the
geometry is invented, every figure caption says so, and the counts drawn obey the same rules the
survey uses (cumulative bands; a 500 m stretch is high with more than five structures inside the
widest band, medium with one to five, low with none), so the picture never contradicts the method.
Deterministic: the same code always draws the same pixels.
"""
import math
import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from shapely.geometry import LineString, Point, Polygon
from shapely.ops import substring

SITE = Path(__file__).resolve().parents[1]
OUT = SITE / "images" / "diagrams"
FONT = SITE / "fonts" / "inter-latin-wght-normal.woff2"
SS = 2                       # supersampling factor
W, H = 1800, 1100            # output size

BG, GRID, TEXT, DIM, LINE = "#ffffff", "#eef1f4", "#172030", "#4b5563", "#c9d0d9"
ORANGE, RED, AMBER, TEAL, GREY = "#e8672f", "#dc2626", "#f59e0b", "#0f766e", "#cbd2db"
PANEL = "#f4f6f8"


def font(size, weight=500):
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
    def __init__(self, h=H):
        self.h = h
        self.im = Image.new("RGBA", (W * SS, h * SS), rgba(BG))
        self.d = ImageDraw.Draw(self.im, "RGBA")

    def poly(self, pts, fill=None, outline=None, width=2):
        pts = [(x * SS, y * SS) for x, y in pts]
        if fill:
            layer = Image.new("RGBA", self.im.size, (0, 0, 0, 0))
            ImageDraw.Draw(layer).polygon(pts, fill=fill)
            self.im.alpha_composite(layer)
        if outline:
            self.d.line(pts + [pts[0]], fill=outline, width=width * SS, joint="curve")

    def line(self, pts, color, width=2):
        self.d.line([(x * SS, y * SS) for x, y in pts], fill=color, width=width * SS, joint="curve")

    def dashed(self, pts, color, width=2, dash=14, gap=10):
        segs, carry, on = [], 0.0, True
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            L = math.hypot(x1 - x0, y1 - y0)
            pos = 0.0
            while pos < L:
                step = min((dash if on else gap) - carry, L - pos)
                a = (x0 + (x1 - x0) * pos / L, y0 + (y1 - y0) * pos / L)
                b = (x0 + (x1 - x0) * (pos + step) / L, y0 + (y1 - y0) * (pos + step) / L)
                if on:
                    segs.append((a, b))
                pos += step
                carry += step
                if carry >= (dash if on else gap) - 1e-9:
                    carry, on = 0.0, not on
        for a, b in segs:
            self.line([a, b], color, width)

    def rect(self, box, fill=None, outline=None, width=1, radius=0):
        box = [v * SS for v in box]
        self.d.rounded_rectangle(box, radius=radius * SS, fill=fill, outline=outline, width=width * SS)

    def text(self, xy, s, size=26, weight=500, color=TEXT, anchor="la"):
        self.d.text((xy[0] * SS, xy[1] * SS), s, font=font(size, weight), fill=color, anchor=anchor)

    def square(self, c, color, size=15, ring=None):
        x, y = c
        h = size / 2
        if ring:
            self.rect([x - h - 6, y - h - 6, x + h + 6, y + h + 6], outline=ring, width=4, radius=4)
        self.rect([x - h, y - h, x + h, y + h], fill=color, outline="#ffffff", width=2, radius=2)

    def grid(self, box, step=60):
        x0, y0, x1, y1 = box
        for x in range(x0, x1 + 1, step):
            self.line([(x, y0), (x, y1)], GRID, 1)
        for y in range(y0, y1 + 1, step):
            self.line([(x0, y), (x1, y)], GRID, 1)

    def save(self, path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.im.resize((W, self.h), Image.LANCZOS).convert("RGB").save(path, optimize=True)
        print("wrote", path.relative_to(SITE.parent))


# ---------------------------------------------------------------- corridor
def corridor():
    c = Canvas(900)
    mx0, my0, mx1, my1 = 30, 30, W - 30, c.h - 30
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    c.grid((mx0 + 20, my0 + 20, mx1 - 20, my1 - 20))
    length = 2500.0

    def route_y(x):
        return 230 * math.sin(x / 560.0 + 0.3) + 45 * math.sin(x / 210.0 + 1.1)

    route = LineString([(x, route_y(x)) for x in [i * 10.0 for i in range(400)]])
    route = substring(route, 0, length)          # exactly 2.5 km along the line
    rnd = random.Random(7)
    plan = {0: [(120, 210), (230, -150), (360, 260), (420, -230)],
            1: [(610, 70), (700, -38), (860, 92), (930, 190), (780, -260)],
            2: [(1060, 25), (1100, -62), (1150, 44), (1195, -84), (1240, 16), (1290, 72),
                (1345, -30), (1400, 90), (1420, -150), (1460, 210), (1320, -240)],
            3: [(1640, -85), (1760, 190), (1880, -170), (1950, 260)],
            4: [(2080, 150), (2200, -240), (2300, 200), (2420, -170)]}
    pts, counts = [], {}
    for seg, lst in plan.items():
        counts[seg] = 0
        for ch, off in lst:
            ch += rnd.uniform(-8, 8)
            base = route.interpolate(ch)
            nxt = route.interpolate(min(ch + 5, length))
            dx, dy = nxt.x - base.x, nxt.y - base.y
            n = math.hypot(dx, dy)
            p = Point(base.x - dy / n * off, base.y + dx / n * off)
            d = route.distance(p)
            counts[seg] += d <= 100
            pts.append((p, d))

    # fit the drawing (route, bands, structures) into the map, leaving room for the rating strip
    strip_h = 150
    x0, y0, x1, y1 = route.buffer(100).bounds
    for p, _ in pts:
        x0, y0, x1, y1 = min(x0, p.x), min(y0, p.y), max(x1, p.x), max(y1, p.y)
    y0 -= 60                                   # room for chainage labels
    pad = 60
    bw, bh = (mx1 - mx0 - 2 * pad), (my1 - my0 - 2 * pad - strip_h)
    k = min(bw / (x1 - x0), bh / (y1 - y0))
    ox = mx0 + pad + (bw - (x1 - x0) * k) / 2 - x0 * k
    oy = my0 + pad + (bh - (y1 - y0) * k) / 2 - y0 * k
    to_px = lambda q: (ox + q[0] * k, oy + q[1] * k)

    for dist, color in ((100, rgba(AMBER, 60)), (50, rgba(RED, 55))):
        band = route.buffer(dist, cap_style=2)
        c.poly([to_px(q) for q in band.exterior.coords], fill=color)
        c.line([to_px(q) for q in band.exterior.coords], rgba(AMBER if dist == 100 else RED, 150), 1)
    for p, d in pts:
        c.square(to_px((p.x, p.y)), RED if d <= 50 else AMBER if d <= 100 else TEAL, 20)
    c.line([to_px(q) for q in route.coords], ORANGE, 8)
    for i in range(6):
        q = route.interpolate(i * 500.0)
        x, y = to_px((q.x, q.y))
        c.line([(x, y - 20), (x, y + 20)], TEXT, 4)
        c.text((x, y - 34), f"km {i * 0.5:.1f}", 26, 700, TEXT, anchor="ms")
    s0 = route.interpolate(60)
    lx, ly = to_px((s0.x, s0.y))
    c.text((lx - 20, ly + 50 * k * 0.55), "50 m", 26, 750, RED, anchor="lm")
    c.text((lx - 20, ly + 100 * k * 0.85), "100 m", 26, 750, "#b45309", anchor="lm")
    c.text((mx0 + 34, my0 + 42), "SCHEMATIC", 24, 750, TEAL)

    sy = my1 - pad - 56
    c.text((mx0 + pad, sy - 16), "Encroachment density per 500 m", 26, 700, DIM, anchor="ls")
    ticks = [to_px(route.interpolate(i * 500.0).coords[0])[0] for i in range(6)]
    for seg in range(5):
        n = counts[seg]
        rating, color = ("High", RED) if n > 5 else ("Medium", AMBER) if n >= 1 else ("Low", GREY)
        xa, xb = ticks[seg] + 4, ticks[seg + 1] - 4
        c.rect([xa, sy, xb, sy + 56], fill=color, radius=8)
        c.text(((xa + xb) / 2, sy + 29), f"{rating} · {n}", 28, 750, "#ffffff" if n > 0 else TEXT, anchor="mm")
    c.save(OUT / "corridor.png")


# ---------------------------------------------------------------- area and ring
def area_ring():
    c = Canvas()
    mx0, my0, mx1, my1 = 30, 30, W - 30, c.h - 30
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    c.grid((mx0 + 20, my0 + 20, mx1 - 20, my1 - 20))

    boundary = Polygon([(0, 250), (700, 40), (1500, 0), (2300, 180), (2800, 650), (2850, 1350),
                        (2350, 1900), (1400, 2050), (500, 1850), (60, 1250)])
    ring = boundary.buffer(260, join_style=1)
    setback = boundary.buffer(-70, join_style=2)
    tsf = Polygon([(1850, 650), (2300, 600), (2450, 950), (2150, 1150), (1800, 1020)])
    zone = Polygon([(1800, 1000), (2150, 1150), (2450, 950), (2850, 1450), (3050, 2000),
                    (2800, 2330), (2400, 2250), (2050, 1750)])
    pit = Point(850, 900).buffer(320).union(Point(1100, 980).buffer(260))
    track = [(-620, 1500), (-400, 1440), (-180, 1370), (60, 1330), (330, 1280)]
    rnd = random.Random(11)
    clusters = [((-450, 650), 6, 140, None), ((60, 2150), 6, 170, None), ((3300, 1150), 3, 110, None),
                ((2750, 2150), 5, 130, None), ((2450, 1700), 3, 90, None),
                ((260, 1400), 4, 110, "in"), ((1150, 1880), 3, 90, "in"), ((1300, -150), 4, 80, None),
                ((2720, 330), 4, 70, None), ((3050, 1600), 3, 60, None)]
    pts = []
    for (cx, cy), n, spread, want in clusters:
        for _ in range(n * 3):
            p = Point(cx + rnd.uniform(-spread, spread), cy + rnd.uniform(-spread, spread))
            if tsf.buffer(40).contains(p) or pit.buffer(40).contains(p):
                continue
            if want == "in" and not setback.buffer(40).contains(p):
                continue
            pts.append(p)
            if sum(1 for q in pts if q.distance(Point(cx, cy)) <= spread * 1.5) >= n:
                break
    s_pt = Point(3120, 800)

    xs, ys = [], []
    for g in (ring, zone):
        x0, y0, x1, y1 = g.bounds
        xs += [x0, x1]; ys += [y0, y1]
    xs += [p.x for p in pts] + [x for x, _ in track] + [3500]
    ys += [p.y for p in pts] + [2450]
    gx0, gx1, gy0, gy1 = min(xs), max(xs), min(ys), max(ys)
    pad = 55
    k = min((mx1 - mx0 - 2 * pad) / (gx1 - gx0), (my1 - my0 - 2 * pad) / (gy1 - gy0))
    ox = mx0 + pad + ((mx1 - mx0 - 2 * pad) - (gx1 - gx0) * k) / 2 - gx0 * k
    oy = my0 + pad + ((my1 - my0 - 2 * pad) - (gy1 - gy0) * k) / 2 - gy0 * k
    to_px = lambda p: (ox + p[0] * k, oy + p[1] * k)

    c.poly([to_px(p) for p in ring.exterior.coords], fill=rgba(AMBER, 34))
    c.poly([to_px(p) for p in boundary.exterior.coords], fill=rgba(ORANGE, 26))
    c.poly([to_px(p) for p in zone.exterior.coords], fill=rgba(RED, 40))
    c.dashed([to_px(p) for p in zone.exterior.coords], RED, 3, 14, 9)
    c.dashed([to_px(p) for p in ring.exterior.coords], "#b45309", 3, 16, 10)
    c.dashed([to_px(p) for p in setback.exterior.coords], ORANGE, 2, 8, 8)
    c.line([to_px(p) for p in boundary.exterior.coords], ORANGE, 6)
    c.poly([to_px(p) for p in tsf.exterior.coords], fill=rgba("#6b7280", 170), outline="#4b5563", width=3)
    c.poly([to_px(p) for p in pit.exterior.coords], fill=rgba("#9ca3af", 120), outline="#6b7280", width=2)
    c.text(to_px((2130, 860)), "Tailings", 26, 750, "#ffffff", anchor="mm")
    c.text(to_px((2130, 950)), "facility", 26, 750, "#ffffff", anchor="mm")
    c.text(to_px((950, 930)), "Pit", 28, 750, TEXT, anchor="mm")

    c.dashed([to_px(p) for p in track], "#7c4a1e", 4, 12, 8)
    tx, ty = to_px(track[0])
    c.text((tx, ty + 40), "New track since the last survey", 25, 700, "#7c4a1e")

    for p in pts:
        color = RED if boundary.contains(p) else AMBER if ring.contains(p) else TEAL
        c.square(to_px((p.x, p.y)), color, 19, ring=TEXT if zone.contains(p) else None)

    near = boundary.exterior.interpolate(boundary.exterior.project(s_pt))
    c.dashed([to_px((s_pt.x, s_pt.y)), to_px((near.x, near.y))], TEXT, 3, 8, 6)
    c.square(to_px((s_pt.x, s_pt.y)), TEAL, 19)
    lx, ly = to_px((s_pt.x, s_pt.y))
    c.text((lx + 26, ly - 26), "distance to", 25, 700, TEXT)
    c.text((lx + 26, ly + 6), "the boundary", 25, 700, TEXT)

    bx, by = to_px((650, 150))
    c.text((bx, by), "Your boundary", 28, 800, "#9a3412")
    rx, ry = to_px((150, -60))
    c.text((rx, ry), "Ring around it", 28, 800, "#b45309")
    zx, zy = to_px((2250, 2330))
    c.text((zx - 60, zy + 34), "Zone supplied by your engineers", 27, 800, RED)

    c.text((mx0 + 34, my0 + 42), "SCHEMATIC", 24, 750, TEAL)
    c.save(OUT / "area-ring.png")


if __name__ == "__main__":
    corridor()
    area_ring()
