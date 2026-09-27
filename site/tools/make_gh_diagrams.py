#!/usr/bin/env python3
"""Draw the Ghana transmission-line diagram used on the /gh/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_gh_diagrams.py

Writes site/images/diagrams/gh-line-strips.png: a plan view of two overhead lines drawn to one scale,
each with the protected strip the Volta River Authority (Transmission Line Protection) Regulations 1967
(L.I. 542), as amended by L.I. 1737 (2004), set either side of the centre of the line, as GRIDCo
described them in September 2026: 20 m each side of a 330 kV line and 15 m each side of a 161 kV
line. The only words in the image are the voltages and the widths; the page caption names the
regulations. The structures are invented and coloured the way the register bands them (red inside
the strip, teal outside). Deterministic: the same code always draws the same pixels. Reuses the Canvas
helpers of make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, TEAL, TEXT, W, H, Canvas, rgba

LINE_COLOR = "#e8672f"
K = 7.0                                         # pixels per metre, the same for both lines

# voltage label, protected distance either side of the centre line (m), invented structures as
# (x, metres from the centre line)
ROWS = [
    ("330 kV", 20, [(560, 8), (760, -15), (980, 12), (1230, -6), (1420, 17),
                    (640, 26), (870, -27), (1100, 28), (1330, -25), (1500, -29)]),
    ("161 kV", 15, [(620, -5), (900, 10), (1180, -11), (1440, 4),
                    (560, 20), (760, -21), (1040, 22), (1300, 19), (1500, -23)]),
]
MARGIN = 66                                     # space above and below each strip for outside structures


def tower(c, x, y):
    s = 15
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


def line_strips():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    bx0, bx1 = 480, mx1 - 36
    heights = [2 * half * K + 2 * MARGIN for _, half, _ in ROWS]
    gap = (my1 - my0 - 40 - sum(heights)) / (len(ROWS) - 1)
    top = my0 + 20
    for (label, half, plan), hgt in zip(ROWS, heights):
        axis = top + hgt / 2
        y = lambda m, a=axis: a - m * K
        c.fill([bx0, y(half), bx1, y(-half)], rgba(RED, 58))
        for side in (1, -1):
            c.line([(bx0, y(half * side)), (bx1, y(half * side))], rgba(RED, 190), 3)
        for x, m in plan:
            c.square((x, y(m)), RED if abs(m) <= half else TEAL, 24)
        c.line([(bx0 - 10, axis), (bx1, axis)], LINE_COLOR, 7)
        for tx in (bx0 + 40, (bx0 + bx1) / 2 + 20, bx1 - 40):
            tower(c, tx, axis)
        # labels and the two dimension arrows (one per side) in the left gutter
        c.text((60, axis - 8), label, 58, 800, TEXT, anchor="ls")
        ax = 400
        c.arrow_v(ax, y(half), axis, RED, 5, 12)
        c.arrow_v(ax, axis, y(-half), RED, 5, 12)
        c.line([(ax - 14, y(half)), (bx0, y(half))], rgba(RED, 150), 2)
        c.line([(ax - 14, y(-half)), (bx0, y(-half))], rgba(RED, 150), 2)
        c.line([(ax - 14, axis), (bx0 - 10, axis)], rgba(DIM, 120), 2)
        c.text((60, axis + 52), f"{half} m + {half} m", 44, 700, "#b91c1c", anchor="ls")
        c.line([(60, axis + 4), (340, axis + 4)], rgba(DIM, 90), 2)
        top += hgt + gap
    c.save(OUT / "gh-line-strips.png")


if __name__ == "__main__":
    line_strips()
