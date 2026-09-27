#!/usr/bin/env python3
"""Draw the Kenya wayleave diagram used on the /ke/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_ke_diagrams.py

Writes site/images/diagrams/ke-rap-wayleaves.png: a plan view of two overhead lines drawn to one
scale, each with the wayleave width its published resettlement action plan sets: 40 m (20 m each
side of the centreline) for a 220 kV line and 30 m (15 m each side) for a 132 kV line
(regional/ke.FINAL.md, "Statutory protection strips"). The only words in the image are the voltages
and the widths; the page caption names the source. The structures are invented and coloured the way
the register bands them (red inside the wayleave, teal outside). Deterministic: the same code always
draws the same pixels. Reuses the Canvas helpers of make_mz_diagrams.py and the tower symbol of
make_ng_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, TEAL, TEXT, W, H, Canvas, rgba
from make_ng_diagrams import LINE_COLOR, tower

K = 7.0                                         # pixels per metre, the same for both lines

# voltage label, wayleave width (m), invented structures as (x, metres from the centreline)
ROWS = [
    ("220 kV", 40, [(600, 8), (760, -14), (980, 17), (1230, -5), (1420, 12),
                    (660, 26), (880, -27), (1100, 29), (1330, -25), (1500, 24)]),
    ("132 kV", 30, [(700, -6), (1010, 11), (1320, -9),
                    (620, 20), (820, -21), (1150, 22), (1460, -19)]),
]
MARGIN = 62                                     # space above and below each band for outside structures


def rap_wayleaves():
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
        c.text((60, axis - 8), label, 58, 800, TEXT, anchor="ls")
        ax = 400
        c.arrow_v(ax, y(half), y(-half), RED, 5, 16)
        c.line([(ax - 14, y(half)), (bx0, y(half))], rgba(RED, 150), 2)
        c.line([(ax - 14, y(-half)), (bx0, y(-half))], rgba(RED, 150), 2)
        c.text((60, axis + 52), f"{width} m", 50, 700, "#b91c1c", anchor="ls")
        c.text((60, axis + 104), f"{width // 2} m each side", 34, 600, DIM, anchor="ls")
        c.line([(60, axis + 4), (340, axis + 4)], rgba(DIM, 90), 2)
        top += hgt + gap
    c.save(OUT / "ke-rap-wayleaves.png")


if __name__ == "__main__":
    rap_wayleaves()
