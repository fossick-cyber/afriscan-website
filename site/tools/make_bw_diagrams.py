#!/usr/bin/env python3
"""Draw the Botswana road-reserve diagram used on /bw/power-lines.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_bw_diagrams.py

Writes site/images/diagrams/bw-road-reserve.png: a plan view, drawn to one scale, of a main road in
its reserve with a 66 kV line built along the reserve's edge. It follows the Environmental and
Social Management Framework of 17 May 2024 for BPC's World Bank-financed 66 kV and 33 kV lines
(Table 2-1 and the road-servitude discussion): main-road servitudes are "typically 60 m wide", the
Roads Department wants lines "as close to the edge of the servitude as possible", and so "half of
the transmission line right of way" falls on the land beside the road: 15 m of the 30 m right of way
of a 66 kV line. The permitted use of that right of way is "Grazing. No cultivation or built
infrastructure permitted."

The only words in the image are the widths and the voltage; the page caption carries the legend and
the source. Structures and the field are invented and coloured the way the register bands them (red
inside the right of way, teal outside). Deterministic: the same code always draws the same pixels.
Reuses the Canvas helpers of make_mz_diagrams.py unchanged.
"""
from make_mz_diagrams import BG, DIM, GRID, LINE, OUT, RED, SS, TEAL, TEXT, W, H, Canvas, font, rgba

ORANGE = "#e8672f"
BROWN = "#8a5a2b"
ROAD, RESERVE = "#5b6470", "#d5d9df"
DARK_RED = "#b91c1c"

K = 7.2                         # pixels per metre
TOP_M, BOTTOM_M = 85, -40       # metres shown above and below the road centreline
RESERVE_HALF = 30               # 60 m road reserve
CARRIAGEWAY_HALF = 5
LINE_M = RESERVE_HALF           # the line runs along the reserve edge
ROW_HALF = 15                   # 30 m right of way for a 66 kV line

# invented structures as (x, metres from the road centreline)
STRUCTURES = [(640, 36), (700, 41), (790, 34), (1330, 39), (1420, 33),
              (600, 55), (760, 63), (840, 52), (1100, 71), (1230, 57), (1480, 66), (1360, 78), (930, 80),
              (660, -35), (1050, -36), (1400, -34)]
FIELD = (880, 31.5, 1190, 60)   # x0, metres, x1, metres: a ploughed field that runs into the right of way


def pill(c, xy, s, size, color):
    f = font(size, 800)
    x, y = xy[0] * SS, xy[1] * SS
    l, t, r, b = c.d.textbbox((x, y), s, font=f, anchor="mm")
    pad = 12 * SS
    c.d.rounded_rectangle([l - pad, t - pad * 0.7, r + pad, b + pad * 0.7], radius=8 * SS,
                          fill=rgba("#ffffff", 240))
    c.d.text((x, y), s, font=f, fill=color, anchor="mm")


def road_reserve():
    c = Canvas()
    mx0, my0, mx1, my1 = 24, 24, W - 24, H - 24
    c.rect([mx0, my0, mx1, my1], fill=BG, outline=LINE, width=2, radius=18)
    for x in range(mx0 + 40, mx1 - 20, 64):
        c.line([(x, my0 + 20), (x, my1 - 20)], GRID, 1)
    for y in range(my0 + 40, my1 - 20, 64):
        c.line([(mx0 + 20, y), (mx1 - 20, y)], GRID, 1)

    centre = 50 + TOP_M * K
    y = lambda m: centre - m * K
    bx0, bx1 = 560, mx1 - 2

    # road reserve and carriageway
    c.fill([bx0, y(RESERVE_HALF), bx1, y(-RESERVE_HALF)], rgba(RESERVE, 255))
    for m in (RESERVE_HALF, -RESERVE_HALF):
        c.line([(bx0, y(m)), (bx1, y(m))], rgba(DIM, 230), 3)
    c.fill([bx0, y(CARRIAGEWAY_HALF), bx1, y(-CARRIAGEWAY_HALF)], rgba(ROAD, 255))
    c.dashed_h(bx0, bx1, y(0), rgba("#ffffff", 255), 3, 26, 22)

    # the line's right of way, half in the reserve and half on the land beside it
    c.fill([bx0, y(LINE_M + ROW_HALF), bx1, y(LINE_M - ROW_HALF)], rgba(RED, 52))
    for m in (LINE_M + ROW_HALF, LINE_M - ROW_HALF):
        c.line([(bx0, y(m)), (bx1, y(m))], rgba(RED, 200), 3)

    # a ploughed field that runs into the right of way
    fx0, fm0, fx1, fm1 = FIELD
    c.fill([fx0, y(fm1), fx1, y(fm0)], rgba(BROWN, 60))
    for i, x in enumerate(range(fx0 + 12, fx1, 22)):
        c.line([(x, y(fm1) + 4), (x, y(fm0) - 4)], rgba(BROWN, 150), 2)
    c.rect([fx0, y(fm1), fx1, y(fm0)], outline=rgba(BROWN, 220), width=3, radius=2)

    for x, m in STRUCTURES:
        inside = LINE_M - ROW_HALF <= m <= LINE_M + ROW_HALF
        c.square((x, y(m)), RED if inside else TEAL)

    # the 66 kV line and its poles, about one span apart
    c.line([(bx0 - 10, y(LINE_M)), (bx1, y(LINE_M))], ORANGE, 7)
    for px in range(bx0 + 70, bx1, 330):
        c.d.ellipse([(px - 11) * SS, (y(LINE_M) - 11) * SS, (px + 11) * SS, (y(LINE_M) + 11) * SS],
                    fill=rgba("#ffffff", 255), outline=rgba(TEXT, 255), width=3 * SS)
    pill(c, (bx1 - 130, y(LINE_M + ROW_HALF) - 36), "66 kV", 44, TEXT)

    # dimension arrows in the left gutter
    for gx, top, bottom, color, label in (
            (96, RESERVE_HALF, -RESERVE_HALF, DIM, "60 m"),
            (262, LINE_M + ROW_HALF, LINE_M - ROW_HALF, RED, "30 m"),
            (420, LINE_M + ROW_HALF, LINE_M, DARK_RED, "15 m")):
        c.arrow_v(gx, y(top), y(bottom), color, 5, 16)
        c.line([(gx - 14, y(top)), (bx0, y(top))], rgba(color, 120), 2)
        c.line([(gx - 14, y(bottom)), (bx0, y(bottom))], rgba(color, 120), 2)
        pill(c, (gx + 72, (y(top) + y(bottom)) / 2), label, 44, color)
    c.save(OUT / "bw-road-reserve.png")


if __name__ == "__main__":
    road_reserve()
