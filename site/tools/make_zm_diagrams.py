#!/usr/bin/env python3
"""Draw the Zambia schematics used on the /zm/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_zm_diagrams.py

Writes to site/images/diagrams/:
  zm-wayleave-widths.png  Electricity (Wayleave and Clearances) Regulations 2026 (SI No. 2 of 2026): the
                          minimum wayleave for one line of 132, 330 and 400 kV (Table 2.1: 32, 48 and 50 m)
                          as a red band, and the wider strip in a recognised forestry area (Table 2.4: 72, 78
                          and 80 m) as dashed green lines, all drawn to one scale and centred on the line.
  zm-mining-consent.png   Minerals Regulation Commission Act 2024 s.35: the areas where a mining-right holder
                          needs written consent before working, around features a structure register maps:
                          180 m around an occupied house, 100 m from a railway track, 90 m around a dam and
                          45 m around cleared or cropped land. Plan view, drawn to one scale.
The only words in the images are the voltages and the distances; the page captions name the regulations
and carry the legend, so the images stay readable on a phone. Geometry and structures are invented.
Deterministic: the same code always draws the same pixels. Reuses the Canvas helpers of
make_za_diagrams.py unchanged.
"""
import math

from make_za_diagrams import BG, LINE, RED, SS, TEAL, TEXT, Canvas, rgba

ORANGE = "#e8672f"
GREEN = "#15803d"
BLUE = "#2563eb"
RAIL = "#374151"


def tower(c, x, y):
    s = 13
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


def dashed_h(c, x0, x1, y, color, width=4):
    c.dashed([(x0, y), (x1, y)], color, width, 22, 14)


# ---------------------------------------------------------------- SI 2 of 2026 wayleave widths
# voltage, Table 2.1 width (m), Table 2.4 forestry width (m), invented structures as (x, metres from the axis)
ROWS = [
    ("132 kV", 32, 72, [(600, 8), (760, -13), (1010, 11), (1230, -6), (1420, 14),
                        (660, 25), (880, -29), (1120, 33), (1330, -24), (1500, -31)]),
    ("330 kV", 48, 78, [(640, -12), (900, 19), (1180, -21), (1400, 9),
                        (560, 30), (790, -33), (1060, 37), (1290, 29), (1480, -36)]),
    ("400 kV", 50, 80, [(700, 17), (980, -9), (1270, 22),
                        (600, -31), (840, 34), (1130, -38), (1360, -30), (1500, 35)]),
]


def wayleave_widths():
    W, H = 1600, 1120
    c = Canvas(W, H)
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    c.grid((mx0 + 20, my0 + 20, mx1 - 20, my1 - 20))

    k = 3.25                                    # pixels per metre, the same for all three lines
    margin = 18
    bx0, bx1 = 500, mx1 - 36
    heights = [fw * k + 2 * margin for _, _, fw, _ in ROWS]
    gap = (my1 - my0 - 40 - sum(heights)) / (len(ROWS) - 1)
    top = my0 + 20
    for (label, width, forest, plan), hgt in zip(ROWS, heights):
        axis = top + hgt / 2
        y = lambda m, a=axis: a - m * k
        half, fhalf = width / 2, forest / 2
        c.fill_rect([bx0, y(fhalf), bx1, y(-fhalf)], rgba(GREEN, 22))
        c.fill_rect([bx0, y(half), bx1, y(-half)], rgba(RED, 60))
        for side in (1, -1):
            c.line([(bx0, y(half * side)), (bx1, y(half * side))], rgba(RED, 200), 3)
            dashed_h(c, bx0, bx1, y(fhalf * side), rgba(GREEN, 235), 4)
        for x, m in plan:
            c.square((x, y(m)), RED if abs(m) <= half else TEAL, 22)
        c.line([(bx0 - 10, axis), (bx1, axis)], ORANGE, 7)
        for tx in (bx0 + 40, (bx0 + bx1) / 2 + 20, bx1 - 40):
            tower(c, tx, axis)
        # voltage and widths in the left gutter, with the two dimension arrows
        c.text((52, axis - 18), label, 54, 800, TEXT, anchor="ls")
        c.text((52, axis + 34), f"{width} m", 44, 800, "#b91c1c", anchor="ls")
        c.text((200, axis + 34), f"{forest} m", 44, 800, GREEN, anchor="ls")
        for ax, h, col in ((390, half, "#b91c1c"), (450, fhalf, GREEN)):
            c.line([(ax, y(h)), (ax, y(-h))], col, 5)
            for yy, s in ((y(h), 1), (y(-h), -1)):
                c.d.polygon([(ax * SS, yy * SS), ((ax - 7) * SS, (yy + s * 14) * SS), ((ax + 7) * SS, (yy + s * 14) * SS)],
                            fill=col)
            c.line([(ax - 12, y(h)), (bx0, y(h))], rgba(col, 120), 2)
            c.line([(ax - 12, y(-h)), (bx0, y(-h))], rgba(col, 120), 2)
        top += hgt + gap
    c.save("zm-wayleave-widths.png")


# ---------------------------------------------------------------- MRC Act 2024 s.35 consent areas
def mining_consent():
    W, H = 1600, 1000
    c = Canvas(W, H)
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    c.grid((mx0 + 20, my0 + 20, mx1 - 20, my1 - 20))

    k = 0.95                                    # pixels per metre, one scale for every distance
    # licence area boundary (orange outline)
    lx0, ly0, lx1, ly1 = 90, 90, 1510, 910
    c.fill_rect([lx0, ly0, lx1, ly1], rgba(ORANGE, 14))

    # railway track along the bottom, with its 100 m band
    rail_y = 800
    c.fill_rect([mx0 + 2, rail_y - 100 * k, mx1 - 2, rail_y + 100 * k if rail_y + 100 * k < my1 else my1 - 2],
                rgba("#6b7280", 40))
    for yy in (rail_y - 100 * k, rail_y + 100 * k):
        c.dashed([(mx0 + 2, yy), (mx1 - 2, yy)], rgba(RAIL, 230), 4, 20, 12)
    c.line([(mx0 + 2, rail_y), (mx1 - 2, rail_y)], RAIL, 6)
    for x in range(mx0 + 20, mx1 - 10, 34):
        c.line([(x, rail_y - 10), (x, rail_y + 10)], RAIL, 3)

    # cropped field with its 45 m band (upper right)
    fx0, fy0, fx1, fy1 = 1060, 150, 1400, 360
    b = 45 * k
    c.fill_rect([fx0 - b, fy0 - b, fx1 + b, fy1 + b], rgba(GREEN, 34))
    c.dashed([(fx0 - b, fy0 - b), (fx1 + b, fy0 - b), (fx1 + b, fy1 + b), (fx0 - b, fy1 + b), (fx0 - b, fy0 - b)],
             rgba(GREEN, 235), 4, 18, 12)
    c.fill_rect([fx0, fy0, fx1, fy1], rgba("#84cc16", 110))
    for yy in range(int(fy0) + 16, int(fy1), 22):
        c.line([(fx0 + 8, yy), (fx1 - 8, yy)], rgba("#4d7c0f", 150), 2)

    # dam with its 90 m ring (centre right)
    dam = (1180, 560)
    c.fill_circle(dam, 60 + 90 * k, rgba(BLUE, 30))
    c.dashed_circle(dam, 60 + 90 * k, rgba(BLUE, 235), 4)
    c.fill_circle(dam, 60, rgba(BLUE, 150))

    # village: occupied houses, each with a 180 m ring (drawn as their union)
    houses = [(330, 330), (395, 290), (430, 380), (300, 420), (470, 310), (360, 470)]
    r = 180 * k
    for hx, hy in houses:
        c.fill_circle((hx, hy), r, rgba(RED, 14))
    for hx, hy in houses:
        c.dashed_circle((hx, hy), r, rgba(RED, 170), 3)
    for hx, hy in houses:
        c.square((hx, hy), RED, 24)

    # a planned working area in the open ground, clear of every band
    wx0, wy0, wx1, wy1 = 700, 430, 900, 610
    c.fill_rect([wx0, wy0, wx1, wy1], rgba(ORANGE, 70))
    c.rect([wx0, wy0, wx1, wy1], outline=ORANGE, width=4, radius=4)

    # outside the bands: structures further away (teal)
    for p in ((760, 240), (900, 190), (620, 650), (1460, 640), (960, 660)):
        c.square(p, TEAL, 22)

    c.rect([lx0, ly0, lx1, ly1], outline=ORANGE, width=6, radius=6)

    # distance labels on white pills
    hx, hy = houses[3]
    c.line([(hx, hy), (hx - r * math.cos(0.5), hy + r * math.sin(0.5))], rgba("#b91c1c", 255), 4)
    c.label((hx - r * 0.55 * math.cos(0.5) - 10, hy + r * 0.55 * math.sin(0.5) + 30), "180 m", 34, "#b91c1c")
    c.label((1300, rail_y - 100 * k / 2), "100 m", 34, RAIL)
    c.line([(dam[0] + 60, dam[1]), (dam[0] + 60 + 90 * k, dam[1])], rgba(BLUE, 255), 4)
    c.label((dam[0] + 60 + 90 * k / 2, dam[1] - 34), "90 m", 34, BLUE)
    c.label((fx1 + b / 2 + 6, fy1 + b + 30), "45 m", 34, GREEN)
    c.save("zm-mining-consent.png")


if __name__ == "__main__":
    wayleave_widths()
    mining_consent()
