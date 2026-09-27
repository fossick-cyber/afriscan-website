#!/usr/bin/env python3
"""Draw the Tanzania reserve diagram used on the /tz/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_tz_diagrams.py

Writes site/images/diagrams/tz-reserves.png: a plan view, drawn to one scale, of a trunk road and a
railway, each with the reserve Tanzanian law sets for it: 30 m either side of the centre of the road
for trunk and regional roads (Roads Management Regulations, 2009, as quoted by the High Court in 2022)
and 30 m either side of the track centre line (Railways Act, Cap. 170, s.3). The only words in the
image are the two row names and the widths; the page caption names the instruments. The carriageway
and track widths are illustrative, and the structures are invented and coloured the way the register
bands them (red inside the reserve, teal outside). Deterministic: the same code always draws the same
pixels. Reuses the Canvas helpers of make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, TEAL, TEXT, W, H, Canvas, rgba

K = 4.6                                         # pixels per metre, the same for both rows
HALF = 30                                       # reserve half-width in metres, both rows
MARGIN = 70                                     # space above and below each reserve for outside structures
ROAD, VERGE = "#5b6573", "#ffffff"
BALLAST, RAIL = "#a8a29e", "#334155"

# row name, invented structures as (x, metres from the centre line)
ROWS = [
    ("Trunk road", [(560, 12), (640, -22), (760, 26), (905, -9), (1010, 18), (1150, -27), (1300, 8),
                    (1420, -16), (1500, 23),
                    (600, 36), (820, -37), (1080, 38), (1240, -36), (1460, 37)]),
    ("Railway", [(620, 21), (840, -14), (990, 28), (1210, -24), (1380, 11),
                 (580, -37), (720, 37), (930, -38), (1120, 36), (1330, -37), (1490, 38)]),
]


def road(c, x0, x1, axis):
    c.fill([x0, axis - 4 * K, x1, axis + 4 * K], rgba(ROAD, 255))
    c.dashed_h(x0 + 10, x1, axis, VERGE, 3, 26, 18)


def railway(c, x0, x1, axis):
    c.fill([x0, axis - 2.2 * K, x1, axis + 2.2 * K], rgba(BALLAST, 255))
    for x in range(int(x0) + 6, int(x1), 14):
        c.line([(x, axis - 1.8 * K), (x, axis + 1.8 * K)], rgba("#78716c", 255), 3)
    for off in (-0.75, 0.75):
        c.line([(x0, axis + off * K), (x1, axis + off * K)], RAIL, 3)


def reserves():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    bx0, bx1 = 480, mx1 - 36
    hgt = 2 * HALF * K + 2 * MARGIN
    gap = (my1 - my0 - 40 - 2 * hgt)
    top = my0 + 20
    for (label, plan), draw in zip(ROWS, (road, railway)):
        axis = top + hgt / 2
        y = lambda m, a=axis: a - m * K
        c.fill([bx0, y(HALF), bx1, y(-HALF)], rgba(RED, 46))
        for side in (1, -1):
            c.line([(bx0, y(HALF * side)), (bx1, y(HALF * side))], rgba(RED, 190), 3)
        draw(c, bx0 - 10, bx1, axis)
        for x, m in plan:
            c.square((x, y(m)), RED if abs(m) <= HALF else TEAL, 24)
        # row name and the two 30 m dimensions in the left gutter
        c.text((60, axis + 18), label, 50, 800, TEXT, anchor="ls")
        ax = 400
        for side in (1, -1):
            c.arrow_v(ax, y(HALF * side), axis, RED, 5, 14)
            c.line([(ax - 14, y(HALF * side)), (bx0, y(HALF * side))], rgba(RED, 150), 2)
            c.text((ax - 18, (y(HALF * side) + axis) / 2 + 16), "30 m", 40, 700, "#b91c1c", anchor="rs")
        c.line([(ax - 14, axis), (bx0 - 10, axis)], rgba(DIM, 120), 2)
        top += hgt + gap
    c.save(OUT / "tz-reserves.png")


if __name__ == "__main__":
    reserves()
