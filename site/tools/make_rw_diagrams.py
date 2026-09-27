#!/usr/bin/env python3
"""Draw the Rwanda right-of-way diagram used on the /rw/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_rw_diagrams.py

Writes site/images/diagrams/rw-rura-row.png: a plan view of four overhead lines drawn to one scale,
each with the minimum horizontal right of way in Schedule I of RURA's Guidelines
N°01/GL/EL-EWS/RURA/2015 on Right-of-Way for Power Lines, with "the power lines ... centered in the
Right-of-Ways": 50 m for 400 kV, 30 m for 220 kV, 25 m for 110 kV and 12 m for 15-30 kV lines. The
0.4 kV width (3 m) is left out; the page text gives it. The only words in the image are the voltages
and the widths; the page caption names the guideline. The structures are invented and coloured the
way the register bands them (red inside the right of way, teal outside). Deterministic: the same
code always draws the same pixels. Reuses the Canvas helpers of make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, TEAL, TEXT, W, H, Canvas, rgba

LINE_COLOR = "#e8672f"
K = 4.4                                         # pixels per metre, the same for all four lines
MARGIN = 46                                     # space above and below each band for outside structures
MIN_ROW = 150                                   # room for the two labels on the narrow rows

# voltage label, right-of-way width (m), invented structures as (x, metres from the axis)
ROWS = [
    ("400 kV", 50, [(600, 11), (760, -19), (980, 16), (1210, -8), (1400, 19),
                    (660, 30), (880, -31), (1090, 31), (1320, -30)]),
    ("220 kV", 30, [(640, -9), (930, 10), (1260, -10),
                    (590, 22), (800, -23), (1120, 24), (1440, 21)]),
    ("110 kV", 25, [(700, 7), (1050, -8), (1350, 8),
                    (620, 19), (870, -20), (1180, 21), (1470, -18)]),
    ("15–30 kV", 12, [(780, 3), (1230, -3),
                      (640, 12), (930, -13), (1080, 13), (1360, -12), (1480, 12)]),
]


def pole(c, x, y):
    s = 12
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


def rura_row():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    bx0, bx1 = 500, mx1 - 36
    heights = [max(w * K + 2 * MARGIN, MIN_ROW) for _, w, _ in ROWS]
    gap = (my1 - my0 - 40 - sum(heights)) / (len(ROWS) - 1)
    top = my0 + 20
    for (label, width, plan), hgt in zip(ROWS, heights):
        axis = top + hgt / 2
        half = width / 2
        y = lambda m, a=axis: a - m * K
        c.fill([bx0, y(half), bx1, y(-half)], rgba(RED, 58))
        for side in (1, -1):
            c.line([(bx0, y(half * side)), (bx1, y(half * side))], rgba(RED, 190), 3)
        for x, m in plan:
            c.square((x, y(m)), RED if abs(m) <= half else TEAL, 22)
        c.line([(bx0 - 10, axis), (bx1, axis)], LINE_COLOR, 6)
        for tx in (bx0 + 40, (bx0 + bx1) / 2 + 20, bx1 - 40):
            pole(c, tx, axis)
        c.text((56, axis - 8), label, 50, 800, TEXT, anchor="ls")
        ax = 430
        c.arrow_v(ax, y(half), y(-half), RED, 5, 14 if width > 20 else 10)
        c.line([(ax - 14, y(half)), (bx0, y(half))], rgba(RED, 150), 2)
        c.line([(ax - 14, y(-half)), (bx0, y(-half))], rgba(RED, 150), 2)
        c.text((56, axis + 48), f"{width} m", 44, 700, "#b91c1c", anchor="ls")
        c.line([(56, axis + 4), (380, axis + 4)], rgba(DIM, 90), 2)
        top += hgt + gap
    c.save(OUT / "rw-rura-row.png")


if __name__ == "__main__":
    rura_row()
