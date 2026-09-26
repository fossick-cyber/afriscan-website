#!/usr/bin/env python3
"""Draw the Nigeria right-of-way diagram used on the /ng/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_ng_diagrams.py

Writes site/images/diagrams/ng-nesis-row.png: a plan view of three overhead lines drawn to one
scale, each with the right of way NESIS Regulations 2015 Table 3.1 sets for its voltage, "divided
equally from the centre of the line on either side": 50 m for 330 kV, 30 m for 132 kV and 11 m for
33 kV (11 kV lines also take 11 m). The only words in the image are the voltages and the widths; the
page caption names the regulation. The structures are invented and coloured the way the register
bands them (red inside the right of way, teal outside). Deterministic: the same code always draws
the same pixels. Reuses the Canvas helpers of make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, TEAL, TEXT, W, H, Canvas, rgba

LINE_COLOR = "#e8672f"
K = 5.2                                         # pixels per metre, the same for all three lines

# voltage label, right-of-way width (m), invented structures as (x, metres from the axis)
ROWS = [
    ("330 kV", 50, [(560, 9), (700, -19), (905, 14), (1190, -7), (1380, 21),
                    (620, 31), (820, -32), (1040, 34), (1270, -29), (1480, -33)]),
    ("132 kV", 30, [(650, -6), (980, 11), (1300, -12),
                    (590, 21), (780, -22), (1120, 23), (1440, 20)]),
    ("33 kV", 11, [(760, 3), (1210, -4),
                   (640, 11), (900, -12), (1080, 13), (1350, -11), (1500, 12)]),
]
MARGIN = 58                                     # space above and below each band for outside structures


def tower(c, x, y):
    s = 14
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


def nesis_row():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    bx0, bx1 = 480, mx1 - 36
    heights = [w * K + 2 * MARGIN for _, w, _ in ROWS]
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
            c.square((x, y(m)), RED if abs(m) <= half else TEAL, 24)
        c.line([(bx0 - 10, axis), (bx1, axis)], LINE_COLOR, 7)
        for tx in (bx0 + 40, (bx0 + bx1) / 2 + 20, bx1 - 40):
            tower(c, tx, axis)
        # labels and the dimension arrow in the left gutter
        c.text((60, axis - 8), label, 58, 800, TEXT, anchor="ls")
        ax = 400
        c.arrow_v(ax, y(half), y(-half), RED, 5, 16 if width > 20 else 12)
        c.line([(ax - 14, y(half)), (bx0, y(half))], rgba(RED, 150), 2)
        c.line([(ax - 14, y(-half)), (bx0, y(-half))], rgba(RED, 150), 2)
        c.text((60, axis + 52), f"{width} m", 50, 700, "#b91c1c", anchor="ls")
        c.line([(60, axis + 4), (340, axis + 4)], rgba(DIM, 90), 2)
        top += hgt + gap
    c.save(OUT / "ng-nesis-row.png")


if __name__ == "__main__":
    nesis_row()
