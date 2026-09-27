#!/usr/bin/env python3
"""Draw the Uganda corridor diagrams used on the /ug/ pages.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_ug_diagrams.py

Writes two plan views to site/images/diagrams/, each drawn to one scale:

ug-wayleave-corridor.png
    The 60 m corridor in UETCL's resettlement policy framework for a new 400 kV interconnector
    (29 January 2026, Table 3.1): a 10 m right of way acquired outright, centred on the line, and
    25 m of wayleave on each side, where no structures are allowed and vegetation and crops must not
    exceed 2 m. Squares are invented structures (red inside the right of way, amber in the wayleave,
    teal outside); circles are invented tall trees.

ug-pipeline-bands.png
    The Petroleum (RCTMS) Regulations 2016 bands around a buried pipeline: within 6 m of the pipe,
    machines dig only under the licensee's supervision (reg 97), and anyone planning work within
    30 m of the right of way must first establish where the pipe is (reg 92). The grey right of way
    has no number because its width is set line by line; the drawing uses an arbitrary one.

The only words in either image are widths; the page captions name the rules, so the images stay
readable on a phone. Deterministic: the same code always draws the same pixels. Reuses the Canvas
helpers of make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import AMBER, BG, GRID, LINE, OUT, RED, SS, TEAL, TEXT, W, H, Canvas, rgba

LINE_COLOR = "#e8672f"
GREEN = "#15803d"
GREY = "#94a3b8"


def panel(c):
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)
    return mx0, my0, mx1, my1


def tower(c, x, y):
    s = 16
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


def tree(c, x, y, r=17):
    c.d.ellipse([(x - r) * SS, (y - r) * SS, (x + r) * SS, (y + r) * SS], fill=rgba(GREEN, 235),
                outline=(255, 255, 255, 255), width=3 * SS)


def dimension(c, x, y0, y1, color, label, bx0):
    top, bottom = min(y0, y1), max(y0, y1)
    c.arrow_v(x, top, bottom, color, 5, 16 if bottom - top > 60 else 11)
    c.line([(x - 14, top), (bx0, top)], rgba(color, 150), 2)
    c.line([(x - 14, bottom), (bx0, bottom)], rgba(color, 150), 2)
    c.text((x + 14, (top + bottom) / 2 + 18), label, 46, 800, color, anchor="ls")


def wayleave_corridor():
    c = Canvas()
    mx0, my0, mx1, my1 = panel(c)
    k = 12.0                                    # pixels per metre: 30 m each side plus a margin
    axis = (my0 + my1) / 2
    bx0, bx1 = 470, mx1 - 36
    y = lambda m: axis - m * k

    c.fill([bx0, y(30), bx1, y(-30)], rgba(AMBER, 62))
    c.fill([bx0, y(5), bx1, y(-5)], rgba(RED, 70))
    for m in (30, -30):
        c.line([(bx0, y(m)), (bx1, y(m))], rgba("#b45309", 190), 3)
    for m in (5, -5):
        c.line([(bx0, y(m)), (bx1, y(m))], rgba(RED, 200), 3)

    structures = [(760, 2), (1180, -3),
                  (600, 14), (880, -21), (1010, 24), (1300, -12), (1450, 19),
                  (540, 35), (700, -34), (960, 36), (1120, -35), (1380, 34), (1500, -36)]
    for x, m in structures:
        a = abs(m)
        c.square((x, y(m)), RED if a <= 5 else AMBER if a <= 30 else TEAL, 26)
    for x, m in ((650, -17), (1060, 11), (1230, 22), (1420, -24)):
        tree(c, x, y(m))

    c.line([(bx0 - 10, axis), (bx1, axis)], LINE_COLOR, 7)
    for tx in (bx0 + 70, (bx0 + bx1) / 2, bx1 - 70):
        tower(c, tx, axis)

    dimension(c, 315, y(5), y(-5), RED, "10 m", bx0)
    dimension(c, 315, y(30), y(5), "#b45309", "25 m", bx0)
    dimension(c, 315, y(-5), y(-30), "#b45309", "25 m", bx0)
    dimension(c, 100, y(30), y(-30), TEXT, "60 m", bx0)
    c.save(OUT / "ug-wayleave-corridor.png")


def pipeline_bands():
    c = Canvas()
    mx0, my0, mx1, my1 = panel(c)
    k = 8.3                                     # pixels per metre
    axis = (my0 + my1) / 2
    bx0, bx1 = 470, mx1 - 36
    y = lambda m: axis - m * k
    row = 15                                    # arbitrary half-width of the right of way (unlabelled)

    c.fill([bx0, y(row + 30), bx1, y(-row - 30)], rgba(AMBER, 55))
    c.fill([bx0, y(row), bx1, y(-row)], rgba(GREY, 70))
    c.fill([bx0, y(6), bx1, y(-6)], rgba(RED, 70))
    for m in (row + 30, -row - 30):
        c.line([(bx0, y(m)), (bx1, y(m))], rgba("#b45309", 190), 3)
    for m in (row, -row):
        c.dashed_h(bx0, bx1, y(m), rgba("#475569", 230), 4)
    for m in (6, -6):
        c.line([(bx0, y(m)), (bx1, y(m))], rgba(RED, 200), 3)

    structures = [(820, 4),
                  (640, -3), (1260, 18),
                  (560, 22), (760, -27), (930, 33), (1090, -19), (1220, 41), (1360, -38), (1480, 26),
                  (700, 51), (1010, -52), (1420, 50)]
    for x, m in structures:
        a = abs(m)
        c.square((x, y(m)), RED if a <= 6 else AMBER if a <= row + 30 else TEAL, 26)

    c.dashed_h(bx0 - 10, bx1, axis, LINE_COLOR, 8, dash=34, gap=12)

    dimension(c, 315, y(6), axis, RED, "6 m", bx0)
    dimension(c, 315, y(row + 30), y(row), "#b45309", "30 m", bx0)
    dimension(c, 315, y(-row), y(-row - 30), "#b45309", "30 m", bx0)
    c.save(OUT / "ug-pipeline-bands.png")


if __name__ == "__main__":
    wayleave_corridor()
    pipeline_bands()
