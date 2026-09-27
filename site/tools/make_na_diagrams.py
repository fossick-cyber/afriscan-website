#!/usr/bin/env python3
"""Draw the Namibia wayleave diagram used on /na/power-lines.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_na_diagrams.py

Writes site/images/diagrams/na-wayleave.png: a plan view of a 400 kV line drawn to scale, with the
servitude described in a published resettlement framework for a new Namibian 400 kV line (Nov 2023):
80 m wide, 40 m either side of the line, narrowing to 25 m either side through a densely populated
stretch, with a 12 m strip along the line cleared of vegetation for a service road. The only words in
the image are the three widths; the page caption says where they come from and that widths are set
in each wayleave agreement. The structures are invented and coloured the way the register bands them
(red inside the servitude, teal outside). Deterministic: the same code always draws the same pixels.
Reuses the Canvas helpers of make_mz_diagrams.py unchanged.
"""
from PIL import Image, ImageDraw

from make_mz_diagrams import BG, DIM, GRID, LINE, ORANGE, OUT, RED, SS, TEAL, TEXT, W, H, Canvas, rgba

K = 5.6                                          # pixels per metre
X0, X1 = 330, 1360                               # the servitude runs between the two dimension gutters
TAPER = (930, 1010)                              # where the 40 m half-width narrows to 25 m

# invented structures as (x, metres from the line); positive is above the line
PLAN = [(420, 55), (505, -62), (610, 33), (760, -51), (880, 68),
        (1060, 21), (1110, -18), (1190, 31), (1250, -34), (1300, 22),
        (1080, 41), (1150, -45), (1230, 52), (1330, -39), (1030, -58), (1280, 63), (700, 74)]
TOWERS = [390, 690, 990, 1290]


def half_width(x):
    if x <= TAPER[0]:
        return 40.0
    if x >= TAPER[1]:
        return 25.0
    return 40.0 - 15.0 * (x - TAPER[0]) / (TAPER[1] - TAPER[0])


def poly(c, pts, color):
    layer = Image.new("RGBA", c.im.size, (0, 0, 0, 0))
    ImageDraw.Draw(layer).polygon([(x * SS, y * SS) for x, y in pts], fill=color)
    c.im.alpha_composite(layer)


def tower(c, x, y):
    s = 13
    c.rect([x - s, y - s, x + s, y + s], fill="#ffffff", outline=TEXT, width=3, radius=2)
    c.line([(x - s, y - s), (x + s, y + s)], TEXT, 2)
    c.line([(x - s, y + s), (x + s, y - s)], TEXT, 2)


def wayleave():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    axis = (my0 + my1) / 2
    y = lambda m: axis - m * K
    edge = [X0, TAPER[0], TAPER[1], X1]
    top = [(x, y(half_width(x))) for x in edge]
    bottom = [(x, y(-half_width(x))) for x in reversed(edge)]
    poly(c, top + bottom, rgba(RED, 52))
    for side in (1, -1):
        c.line([(x, y(side * half_width(x))) for x in edge], rgba(RED, 200), 3)

    # the 12 m cleared strip along the line
    c.fill([X0, y(6), X1, y(-6)], rgba(DIM, 46))
    c.dashed_h(X0, X1, y(6), rgba(DIM, 200), 3, 16, 10)
    c.dashed_h(X0, X1, y(-6), rgba(DIM, 200), 3, 16, 10)

    for x, m in PLAN:
        c.square((x, y(m)), RED if abs(m) <= half_width(x) else TEAL, 24)
    c.line([(X0 - 10, axis), (X1 + 10, axis)], ORANGE, 7)
    for tx in TOWERS:
        tower(c, tx, axis)

    # left gutter: the 80 m servitude and the 12 m strip
    gx = 92
    c.arrow_v(gx, y(40), y(-40), RED, 5, 18)
    c.line([(gx - 16, y(40)), (X0, y(40))], rgba(RED, 140), 2)
    c.line([(gx - 16, y(-40)), (X0, y(-40))], rgba(RED, 140), 2)
    c.text((gx + 16, y(40) - 14), "80 m", 54, 800, "#b91c1c", anchor="lb")
    sx = 250
    c.arrow_v(sx, y(6), y(-6), DIM, 4, 11)
    c.line([(sx - 12, y(6)), (X0, y(6))], rgba(DIM, 140), 2)
    c.line([(sx - 12, y(-6)), (X0, y(-6))], rgba(DIM, 140), 2)
    c.text((sx - 26, y(-6) + 20), "12 m", 44, 800, DIM, anchor="la")

    # right gutter: the narrower servitude through the settled stretch
    rx = 1420
    c.arrow_v(rx, y(25), y(-25), RED, 5, 16)
    c.line([(X1, y(25)), (rx + 14, y(25))], rgba(RED, 140), 2)
    c.line([(X1, y(-25)), (rx + 14, y(-25))], rgba(RED, 140), 2)
    c.text((rx + 14, y(25) - 14), "50 m", 50, 800, "#b91c1c", anchor="lb")
    c.save(OUT / "na-wayleave.png")


if __name__ == "__main__":
    wayleave()
