#!/usr/bin/env python3
"""Draw the Malawi road-reserve diagram used on the /mw/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_mw_diagrams.py

Writes site/images/diagrams/mw-road-reserves.png: a plan view of three roads drawn to one scale, each
with the road reserve the consolidated Public Roads Act (Cap. 69:02) lists for its class, with the
reserve's centre line down the centre line of the carriageway: 60 m for a main road, 36 m for a
secondary, tertiary or district road, and 18 m for a branch or estate road. The only words in the
image are the widths; the page caption names the road classes and the Act. The carriageways and the
structures are invented and coloured the way the register bands them (red inside the reserve, teal
outside). Deterministic: the same code always draws the same pixels. Reuses the Canvas helpers of
make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, TEAL, TEXT, W, H, Canvas, rgba

ROAD, ROAD_EDGE, MARK = "#9aa3ad", "#6b7480", "#ffffff"
K = 4.6                                         # pixels per metre, the same for all three roads

# width label, reserve width (m), drawn carriageway width (m, schematic only),
# invented structures as (x, metres from the centre line)
ROWS = [
    ("60 m", 60, 8, [(590, 12), (720, -21), (930, 17), (1150, -9), (1330, 25), (1470, -15),
                     (650, 34), (860, -35), (1060, 35), (1260, -34), (1420, 34)]),
    ("36 m", 36, 7, [(700, -9), (1010, 12), (1290, -13), (1450, 8),
                     (610, 23), (820, -24), (1140, 24), (1380, -23)]),
    ("18 m", 18, 5, [(820, 6), (1230, -6),
                     (640, 14), (960, -15), (1100, 15), (1360, -14), (1500, 15)]),
]
MARGIN = 50                                     # space above and below each reserve for outside structures


def road_reserves():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    bx0, bx1 = 480, mx1 - 36
    heights = [w * K + 2 * MARGIN for _, w, _, _ in ROWS]
    gap = (my1 - my0 - 40 - sum(heights)) / (len(ROWS) - 1)
    top = my0 + 20
    for (label, width, carriage, plan), hgt in zip(ROWS, heights):
        axis = top + hgt / 2
        half = width / 2
        y = lambda m, a=axis: a - m * K
        c.fill([bx0, y(half), bx1, y(-half)], rgba(RED, 58))
        for side in (1, -1):
            c.line([(bx0, y(half * side)), (bx1, y(half * side))], rgba(RED, 190), 3)
        # the carriageway, with its centre line: the reserve is centred on it
        c.fill([bx0 - 10, y(carriage / 2), bx1, y(-carriage / 2)], rgba(ROAD, 255))
        for side in (1, -1):
            c.line([(bx0 - 10, y(carriage / 2 * side)), (bx1, y(carriage / 2 * side))], ROAD_EDGE, 2)
        c.dashed_h(bx0 - 10, bx1, axis, MARK, 3, 26, 18)
        for x, m in plan:
            c.square((x, y(m)), RED if abs(m) <= half else TEAL, 24)
        # the width and its dimension arrow in the left gutter
        ax = 400
        c.arrow_v(ax, y(half), y(-half), RED, 5, 16 if width > 20 else 12)
        c.line([(ax - 14, y(half)), (bx0, y(half))], rgba(RED, 150), 2)
        c.line([(ax - 14, y(-half)), (bx0, y(-half))], rgba(RED, 150), 2)
        c.text((60, axis + 20), label, 66, 800, "#b91c1c", anchor="ls")
        c.line([(60, axis + 36), (340, axis + 36)], rgba(DIM, 90), 2)
        top += hgt + gap
    c.save(OUT / "mw-road-reserves.png")


if __name__ == "__main__":
    road_reserves()
