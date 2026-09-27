#!/usr/bin/env python3
"""Draw the DR Congo diagrams used on the /cd/ and /cd/fr/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_cd_diagrams.py

Writes, next to the other masters in site/images/diagrams/:

  cd-emprise-ht.png  a plan view of a high-voltage line axis drawn to scale, with the 25 m on each
                     side that Arrêté interministériel n° 0021 du 29 octobre 1993 (art. 1) lists as
                     the line's emprise, an inner 10 m band and an outer 50 m reporting band. Drawn
                     from the axis, which is how the register measures; whether the 25 m runs from
                     the axis or from the outer conductor is an open point the page states.
  cd-art279.png      a plan view of the land a mining right holder plans to occupy, with the 800 m
                     and 1 000 m distances of Code minier art. 279 (as amended in 2018) drawn around
                     it: houses and buildings within 1 000 m, cultivated fields within 800 m.
                     Distance is symmetric, so a house within 1 000 m of the works is one the works
                     come within 1 000 m of.

The only words in the images are the distances, so one image serves the French and English pages;
each page caption names the text and says what the colours mean. Geometry and structures are
invented. Deterministic: the same code always draws the same pixels. Reuses the Canvas helpers of
make_mz_diagrams.py unchanged.
"""
import math

from PIL import Image, ImageDraw
from shapely.geometry import Point, Polygon

from make_mz_diagrams import AMBER, BG, DIM, GRID, LINE, OUT, RED, SS, TEAL, TEXT, W, H, Canvas, rgba

NNBSP = "\u202f"                                # narrow no-break space between a number and its unit
LINE_COLOR = "#e8672f"
GREEN = "#4d7c0f"
PALE_GREEN = "#a8c08a"                           # fields beyond 800 m
BLUE = "#2563eb"


def frame(c):
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)
    return mx0, my0, mx1, my1


def tower(c, x, y):
    s = 14
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


# ---------------------------------------------------------------- 25 m emprise
def emprise_ht():
    c = Canvas()
    mx0, my0, mx1, my1 = frame(c)
    k = 8.0                                     # pixels per metre: 50 m each side fits the panel
    axis = (my0 + my1) / 2
    bx0, bx1 = 520, mx1 - 36
    y = lambda m: axis - m * k

    for side in (1, -1):
        c.fill([bx0, min(y(50 * side), y(25 * side)), bx1, max(y(50 * side), y(25 * side))], rgba(AMBER, 60))
        c.fill([bx0, min(y(25 * side), y(10 * side)), bx1, max(y(25 * side), y(10 * side))], rgba(RED, 48))
        c.fill([bx0, min(y(10 * side), axis), bx1, max(y(10 * side), axis)], rgba(RED, 92))
        c.line([(bx0, y(25 * side)), (bx1, y(25 * side))], rgba(RED, 220), 4)
        c.dashed_h(bx0, bx1, y(10 * side), rgba("#991b1b", 200), 3)
        c.line([(bx0, y(50 * side)), (bx1, y(50 * side))], rgba("#b45309", 170), 3)

    # invented structures, metres from the axis, coloured by band (red within 25 m, amber 25–50 m)
    plan = [(600, 6), (700, -8), (820, 17), (960, -21), (1080, 13), (1230, -15), (1400, 22), (1480, -4),
            (650, 33), (760, -38), (900, 44), (1020, -31), (1160, 40), (1300, -46), (1440, 36)]
    for x, m in plan:
        a = abs(m)
        c.square((x, y(m)), RED if a <= 25 else AMBER if a <= 50 else TEAL)

    c.line([(bx0 - 10, axis), (bx1, axis)], LINE_COLOR, 8)
    for tx in (bx0 + 60, (bx0 + bx1) / 2 + 10, bx1 - 60):
        tower(c, tx, axis)

    # dimension arrows in the left gutter, measured from the axis
    for gx, m, color, label in ((420, 10, "#991b1b", "10 m"), (285, 25, RED, "25 m"), (140, 50, "#b45309", "50 m")):
        c.arrow_v(gx, y(m), axis, color, 5, 18)
        c.line([(gx - 18, axis), (bx0 - 10, axis)], rgba(DIM, 120), 2)
        c.line([(gx - 18, y(m)), (bx0, y(m))], rgba(color, 150), 2)
        c.text((gx + 12, y(m) - 12), label, 50, 800, color, anchor="lb")
    c.save(OUT / "cd-emprise-ht.png")


# ---------------------------------------------------------------- Code minier art. 279
def art279():
    c = Canvas()
    mx0, my0, mx1, my1 = frame(c)
    k = 0.31                                    # pixels per metre
    ox, oy = 700, (my0 + my1) / 2               # centre of the works area on the canvas

    def px(p):
        return (ox + p[0] * k, oy - p[1] * k)

    # the land the holder plans to occupy (metres), and the permit around it
    works = Polygon([(-420, -160), (-120, -300), (300, -250), (430, 20), (260, 240), (-160, 280), (-450, 90)])
    permit = Polygon([(-900, -700), (650, -760), (1150, -120), (820, 820), (-600, 760), (-1050, 150)])
    rings = [(1000, RED, 40), (800, AMBER, 60)]

    def poly(shape, fill=None, outline=None, width=3):
        pts = [px(p) for p in shape.exterior.coords]
        layer = Image.new("RGBA", c.im.size, (0, 0, 0, 0))
        ImageDraw.Draw(layer).polygon([(x * SS, y * SS) for x, y in pts], fill=fill)
        if fill:
            c.im.alpha_composite(layer)
        if outline:
            c.line(pts, outline, width)

    for dist, color, alpha in rings:
        poly(works.buffer(dist, 64), fill=rgba(color, alpha))
    for dist, color, _ in rings:
        poly(works.buffer(dist, 64), outline=rgba(color if color != AMBER else "#b45309", 220), width=4)

    # the permit boundary, dashed
    pts = [px(p) for p in permit.exterior.coords]
    for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
        L = math.hypot(x1 - x0, y1 - y0)
        n = max(1, int(L // 30))
        for i in range(n):
            if i % 2 == 0:
                a, b = i / n, min(1, (i + 0.6) / n)
                c.line([(x0 + (x1 - x0) * a, y0 + (y1 - y0) * a), (x0 + (x1 - x0) * b, y0 + (y1 - y0) * b)], TEXT, 3)

    poly(works, fill=rgba(LINE_COLOR, 235), outline=rgba("#9a3412", 255), width=3)

    # invented fields (rectangles, metres) and houses (points, metres)
    fields = [(-900, 420, 260, 170), (520, -820, 230, 150), (1330, 640, 250, 180), (-1350, -520, 240, 160),
              (-200, -1150, 260, 150)]
    for fx, fy, fw, fh in fields:
        rect = Polygon([(fx, fy), (fx + fw, fy), (fx + fw, fy + fh), (fx, fy + fh)])
        inside = works.distance(rect) <= 800
        x0, y0 = px((fx, fy + fh))
        x1, y1 = px((fx + fw, fy))
        green = GREEN if inside else PALE_GREEN
        c.rect([x0, y0, x1, y1], fill=rgba(green), outline=rgba(green), width=3, radius=4)
        for hx in range(int(x0) + 8, int(x1) - 4, 14):
            c.line([(hx, y0 + 6), (hx, y1 - 6)], rgba("#ffffff", 150), 2)

    houses = [(-700, 380), (-620, 470), (-760, 520), (-540, 350), (700, 420), (820, 330), (760, 520),
              (1180, -380), (1260, -300), (-1250, 80), (-1320, 200), (360, -900), (440, -980),
              (1500, 200), (1560, 320), (-1500, -250), (100, 1250), (-60, 1300), (-1100, -1000)]
    for hx, hy in houses:
        d = works.distance(Point(hx, hy))
        c.square(px((hx, hy)), RED if d <= 1000 else TEAL, 22)

    c.d.ellipse([(px((980, -620))[0] - 13) * SS, (px((980, -620))[1] - 13) * SS,
                 (px((980, -620))[0] + 13) * SS, (px((980, -620))[1] + 13) * SS], fill=BLUE, outline="#ffffff", width=3 * SS)

    # distance labels on the right, with leader lines to each ring
    b = works.bounds
    edge = (b[2], (b[1] + b[3]) / 2 - 40)
    for dist, color, ty in ((800, "#b45309", oy - 70), (1000, RED, oy + 90)):
        tx = edge[0] + dist
        p = px((tx, edge[1]))
        c.line([p, (1335, ty)], rgba(color, 200), 3)
        c.d.ellipse([(p[0] - 7) * SS, (p[1] - 7) * SS, (p[0] + 7) * SS, (p[1] + 7) * SS], fill=color)
        label = f"{dist:,}".replace(",", " ") + NNBSP + "m"
        c.text((1350, ty + 18), label, 58, 800, color, anchor="ls")
    c.save(OUT / "cd-art279.png")


if __name__ == "__main__":
    emprise_ht()
    art279()
