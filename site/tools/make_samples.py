#!/usr/bin/env python3
"""Rebuild the T-9 sample images and register from the app's stored job.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_samples.py

The reviewer placed the sample's marks on a Google satellite basemap in the app. That basemap is
not reproduced anywhere on the site: its tiles were fetched in a way Google's terms do not allow
for published material, and it has no capture date. What this tool draws instead:

  register views   strip maps drawn from the register alone: each reviewer mark by chainage and by
                   signed distance from the route (north side up), over the 50 m and 100 m bands.
                   No imagery and no coordinates. English and Portuguese (-pt) versions.
  route views      the route on a dated Copernicus Sentinel-2 L2A scene (10 m, free and open data,
                   credited "Contains modified Copernicus Sentinel data 2026"). 10 m pixels cannot
                   show individual structures, so no marks are drawn on it: the hero shows the route
                   and its 100 m band, the overview colours each 500 m of route by its rating.

  site/images/samples/*.jpg     masters for the image pipeline (build.py makes AVIF/WebP)
  site/data/samples/t9.json     route facts, 500 m segment ratings and the register excerpt

Read-only against /opt/favhousecheck. The Sentinel-2 windows are read from the public
sentinel-cogs bucket (AWS Open Data) once and kept in site/.cache/s2/.
"""
import json
import math
import re
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFont
from pyproj import Transformer
from shapely.geometry import LineString, Point
from shapely.ops import substring

SITE = Path(__file__).resolve().parent.parent
JOB = Path("/opt/favhousecheck/results/corridor_909497cd")      # manual review job, 59 marks
ROUTE = Path("/opt/favhousecheck/uploads/corridor_909497cd/T-9 Replacement Pipeline Rev 00.kmz")
FONT = SITE / "fonts/inter-latin-wght-normal.woff2"
OUT_IMG = SITE / "images/samples"
OUT_DATA = SITE / "data/samples/t9.json"
S2_CACHE = SITE / ".cache/s2"

BUFFERS = (50, 100)            # the job's buffers; the largest drives the segment rating
SEGMENT = 500
HIGH_OVER = 5                  # zones.py: high = more than 5 inside the largest buffer
UTM = 32736                    # the job's UTM zone (EPSG:32736, 36S); also the Sentinel-2 tiles' CRS

EXCERPT_KM = (5.0, 6.5)        # the register view and table: three full 500 m segments
ACROSS_M = 125                 # register views show 125 m either side of the route

# Sentinel-2 L2A, 2 August 2026, 0.05 % cloud. The route sits where tiles 36KYA and 36KYB overlap.
S2_DATE = "2026-08-02"
S2_SCENES = ["S2B_36KYA_20260802_0_L2A", "S2B_36KYB_20260802_0_L2A"]
S2_URL = "https://sentinel-cogs.s3.us-west-2.amazonaws.com/sentinel-s2-l2a-cogs/36/K/{sq}/2026/8/{id}/TCI.tif"

INK, TEXT, DIM, GRID, PANEL = (13, 17, 23), (23, 32, 48), (75, 85, 99), (238, 241, 244), (244, 246, 248)
ORANGE, RED, AMBER, TEAL, LOWGREY = (232, 103, 47), (220, 38, 38), (245, 158, 11), (15, 118, 110), (203, 210, 219)

to_utm = Transformer.from_crs(4326, UTM, always_xy=True)

STRINGS = {
    "en": {"tag": "REGISTER VIEW · NO IMAGERY", "title": "T-9 route · km {a}–{b}", "north": "north side",
           "south": "south side", "high": "High", "medium": "Medium", "low": "Low", "dec": "."},
    "pt": {"tag": "VISTA DO REGISTO · SEM IMAGENS", "title": "Traçado T-9 · km {a}–{b}", "north": "lado norte",
           "south": "lado sul", "high": "Alta", "medium": "Média", "low": "Baixa", "dec": ","},
}


def font(size, weight=700):
    f = ImageFont.truetype(str(FONT), size)
    f.set_variation_by_axes([weight])
    return f


def km(x, lang, nd=1):
    return f"{x:.{nd}f}".replace(".", STRINGS[lang]["dec"])


# ------------------------------------------------------------------ register
def load_route():
    kml = zipfile.ZipFile(ROUTE).read("doc.kml").decode()
    pts = [tuple(map(float, c.split(",")[:2]))
           for c in re.findall(r"<coordinates>(.*?)</coordinates>", kml, re.S)[0].split()]
    return LineString([to_utm.transform(*p) for p in pts])


def load_marks(line):
    feats = json.loads((JOB / "gis/building.geojson").read_text())["features"]
    marks = []
    for f in feats:
        assert f["properties"]["source"] == "manual", "sample must be reviewer marks only"
        p = Point(*to_utm.transform(*f["geometry"]["coordinates"]))
        chain = line.project(p)
        a, b = line.interpolate(max(0, chain - 5)), line.interpolate(min(line.length, chain + 5))
        c = line.interpolate(chain)
        side = 1 if (b.x - a.x) * (p.y - c.y) - (b.y - a.y) * (p.x - c.x) > 0 else -1   # +1: left of travel
        marks.append({"p": p, "chain": chain, "dist": line.distance(p), "side": side})
    marks.sort(key=lambda m: m["chain"])
    for i, m in enumerate(marks, 1):
        m["id"] = f"R{i:02d}"
        m["band"] = "0–50 m" if m["dist"] <= 50 else "50–100 m" if m["dist"] <= 100 else "beyond 100 m"
    return marks


def segments(line, marks):
    rows, n = [], int(-(-line.length // SEGMENT))
    for i in range(n):
        s = substring(line, i * SEGMENT, min((i + 1) * SEGMENT, line.length))
        c = sum(s.distance(m["p"]) <= max(BUFFERS) for m in marks)   # boundary marks count in both, as zones.py
        rating = "high" if c > HIGH_OVER else "medium" if c >= 1 else "low"
        rows.append({"from_m": i * SEGMENT, "to_m": int(min((i + 1) * SEGMENT, round(line.length))),
                     "count": c, "rating": rating})
    return rows


def band_colour(d):
    return RED if d <= 50 else AMBER if d <= 100 else TEAL


def strip(marks, segs, k0, k1, out_name, lang, width, ppm_x, ex=2.0, labels=True, badge=None, zone=None):
    """Straight-line strip map: chainage left to right, signed distance up (north side, left of travel).

    zone=(a, b): one 500 m segment. The view then runs 100 m past each end and fades everything outside
    the area the rating counts (every point within 100 m of the segment), so the count on the badge
    matches the marks that stand out."""
    L = STRINGS[lang]
    R = max(BUFFERS)
    if zone:
        k0, k1 = zone[0] - R, zone[1] + R
    S = 2
    ppm_y = ppm_x * ex
    k = 1 if labels else 1.9                    # small views are shown about a third as wide: bigger text
    left, right = (118 if labels else 150), 40
    top = 104 if labels else 104
    plot_w = round((k1 - k0) * ppm_x)
    plot_h = round(2 * ACROSS_M * ppm_y)
    seg_h = 58 if labels else 0
    W = left + plot_w + right
    H = top + plot_h + round(56 * k) + seg_h + 24
    im = Image.new("RGB", (W * S, H * S), (255, 255, 255))
    d = ImageDraw.Draw(im, "RGBA")
    X = lambda ch: (left + (ch - k0) * ppm_x) * S
    Y = lambda off: (top + (ACROSS_M - off) * ppm_y) * S
    # panel and grid
    d.rounded_rectangle([10 * S, 10 * S, (W - 10) * S, (H - 10) * S], radius=14 * S, fill=(255, 255, 255),
                        outline=(201, 208, 217), width=2 * S)
    step = 50 if (k1 - k0) <= 600 else 100
    for ch in range(int(k0), int(k1) + 1, step):
        d.line([X(ch), Y(ACROSS_M), X(ch), Y(-ACROSS_M)], fill=GRID + (255,), width=S)
    for off in (-100, -50, 50, 100):
        d.line([X(k0), Y(off), X(k1), Y(off)], fill=GRID + (255,), width=S)
    # bands (the 100 m band includes the 50 m band)
    d.rectangle([X(k0), Y(100), X(k1), Y(-100)], fill=AMBER + (46,))
    d.rectangle([X(k0), Y(50), X(k1), Y(-50)], fill=RED + (40,))
    for off, col in ((100, AMBER), (-100, AMBER), (50, RED), (-50, RED)):
        d.line([X(k0), Y(off), X(k1), Y(off)], fill=col + (200,), width=2 * S)
    d.line([X(k0), Y(0), X(k1), Y(0)], fill=ORANGE + (255,), width=7 * S)
    inside = lambda m: True
    if zone:
        a, b = zone
        ring = ([(X(a), Y(R)), (X(b), Y(R))]
                + [(X(b + R * math.cos(t)), Y(R * math.sin(t))) for t in [math.pi / 2 - i * math.pi / 48 for i in range(49)]]
                + [(X(b), Y(-R)), (X(a), Y(-R))]
                + [(X(a - R * math.cos(t)), Y(-R * math.sin(t))) for t in [math.pi / 2 - i * math.pi / 48 for i in range(49)]])
        fade = Image.new("RGBA", im.size, (255, 255, 255, 0))
        fd = ImageDraw.Draw(fade)
        fd.rectangle([X(k0), Y(ACROSS_M), X(k1), Y(-ACROSS_M)], fill=(255, 255, 255, 170))
        fd.polygon(ring, fill=(255, 255, 255, 0))
        im.paste(Image.alpha_composite(im.convert("RGBA"), fade).convert("RGB"))
        d = ImageDraw.Draw(im, "RGBA")
        pts = ring + [ring[0]]
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            d.line([x0, y0, x1, y1], fill=INK + (170,), width=2 * S)
        inside = lambda m: min(max(m["chain"] - b, a - m["chain"], 0) ** 2 + m["dist"] ** 2, 1e12) <= R * R
    # segment boundaries
    for g in segs:
        for edge in (g["from_m"], g["to_m"]):
            if k0 < edge < k1 and (not zone or edge in zone):
                y = Y(ACROSS_M)
                while y < Y(-ACROSS_M):
                    d.line([X(edge), y, X(edge), min(y + 10 * S, Y(-ACROSS_M))], fill=DIM + (150,), width=2 * S)
                    y += 18 * S
    # distance axis
    fa = font(round(15 * k) * S, 600)
    for off in (100, 50, 0, -50, -100):
        t = f"{abs(off)} m"
        tw = d.textlength(t, font=fa)
        d.text((X(k0) - tw - 12 * S, Y(off) - round(10 * k) * S), t, font=fa, fill=DIM)
    if labels:
        fs = font(13 * S, 650)
        d.text((X(k0) + 8 * S, Y(ACROSS_M) + 5 * S), L["north"].upper(), font=fs, fill=DIM)
        d.text((X(k0) + 8 * S, Y(-ACROSS_M) - 22 * S), L["south"].upper(), font=fs, fill=DIM)
    # marks
    box = (11 if labels else 14) * S
    lab = font(15 * S, 700)
    placed = []
    shown = [m for m in marks if k0 <= m["chain"] <= k1 and m["dist"] <= ACROSS_M]
    for m in shown:
        cx, cy = X(m["chain"]), Y(m["side"] * m["dist"])
        col = band_colour(m["dist"])
        a_ = 255 if inside(m) else 95
        d.rectangle([cx - box, cy - box, cx + box, cy + box], fill=col + (a_,), outline=(255, 255, 255), width=2 * S)
        d.rectangle([cx - box - 2 * S, cy - box - 2 * S, cx + box + 2 * S, cy + box + 2 * S], outline=INK + (150 * a_ // 255,), width=S)
    if labels:
        for m in shown:
            cx, cy = X(m["chain"]), Y(m["side"] * m["dist"])
            tw = d.textlength(m["id"], font=lab)
            bw, bh = tw + 10 * S, 22 * S
            spots = [(cx - bw / 2, cy - box - bh - 4 * S), (cx - bw / 2, cy + box + 4 * S),
                     (cx + box + 5 * S, cy - bh / 2), (cx - box - 5 * S - bw, cy - bh / 2),
                     (cx - bw / 2, cy - box - 2 * bh - 8 * S), (cx - bw / 2, cy + box + bh + 8 * S)]
            boxes = [(X(q["chain"]) - box, Y(q["side"] * q["dist"]) - box, X(q["chain"]) + box,
                      Y(q["side"] * q["dist"]) + box) for q in shown]
            r = None
            for bx, by in spots:
                cand = (bx, by, bx + bw, by + bh)
                clash = lambda q: cand[0] < q[2] and q[0] < cand[2] and cand[1] < q[3] and q[1] < cand[3]
                if not any(clash(q) for q in placed + boxes):
                    r = cand
                    break
            r = r or (spots[0][0], spots[0][1], spots[0][0] + bw, spots[0][1] + bh)
            placed.append(r)
            d.rounded_rectangle(r, radius=5 * S, fill=INK + (230,))
            d.text((r[0] + 5 * S, r[1] + 2 * S), m["id"], font=lab, fill=(255, 255, 255))
    # chainage axis
    fk = font(round(15 * k) * S, 650)
    y_axis = Y(-ACROSS_M) + round(26 * k) * S
    for ch in range(int(k0), int(k1) + 1, 100):
        major = ch % 500 == 0
        d.line([X(ch), Y(-ACROSS_M), X(ch), Y(-ACROSS_M) + (12 if major else 6) * S], fill=TEXT, width=2 * S)
        if major:
            t = f"km {km(ch / 1000, lang)}"
            tw = d.textlength(t, font=fk)
            tx = min(max(X(ch) - tw / 2, X(k0) - 20 * S), X(k1) - tw + 10 * S)
            d.text((tx, y_axis - 6 * S), t, font=fk, fill=TEXT)
    # segment ratings
    if seg_h:
        fr = font(17 * S, 700)
        yb = y_axis + 30 * S
        for g in segs:
            a, b = max(g["from_m"], k0), min(g["to_m"], k1)
            if b - a < 1:
                continue
            fill = {"high": RED, "medium": AMBER, "low": LOWGREY}[g["rating"]]
            d.rounded_rectangle([X(a) + 4 * S, yb, X(b) - 4 * S, yb + 38 * S], radius=6 * S, fill=fill)
            t = f"{L[g['rating']]} · {g['count']}"
            tw = d.textlength(t, font=fr)
            d.text(((X(a) + X(b)) / 2 - tw / 2, yb + 8 * S), t, font=fr,
                   fill=(255, 255, 255) if g["rating"] != "low" else TEXT)
    # badge (top left, so the page's own "Reviewed" badge can sit top right), tag, title
    ft = font(round(15 * k) * S, 750)
    tag_x, tag_y = 28 * S, 26 * S
    if zone and badge:
        n_in = sum(inside(m) for m in marks)
        assert n_in == badge[1], f"{out_name}: {n_in} marks drawn inside the counting area, rating counts {badge[1]}"
    if badge:
        g, n = badge
        fb = font(round(20 * k) * S, 750)
        t = f"{L[g]} · {n}"
        tw = d.textlength(t, font=fb)
        fill = {"high": RED, "medium": AMBER, "low": LOWGREY}[g]
        bh = round(38 * k) * S
        x0, y0 = 24 * S, 20 * S
        d.rounded_rectangle([x0, y0, x0 + tw + 24 * S, y0 + bh], radius=8 * S, fill=fill)
        d.text((x0 + 12 * S, y0 + bh / 2 - fb.size * 0.62), t, font=fb, fill=(255, 255, 255) if g != "low" else TEXT)
        tag_x, tag_y = x0 + tw + 24 * S + 18 * S, y0 + bh / 2 - ft.size * 0.62
    d.text((tag_x, tag_y), L["tag"], font=ft, fill=TEAL)
    if labels:
        d.text((28 * S, 52 * S), L["title"].format(a=km(k0 / 1000, lang), b=km(k1 / 1000, lang)),
               font=font(22 * S, 750), fill=TEXT)
    im = im.resize((W, H), Image.LANCZOS)
    OUT_IMG.mkdir(parents=True, exist_ok=True)
    im.save(OUT_IMG / out_name, "JPEG", quality=92, optimize=True, progressive=True)
    return [W, H], [m["id"] for m in shown]


# ------------------------------------------------------------------ Sentinel-2
def s2_window(left, bottom, right, top):
    """True-colour (TCI) window at 10 m from both overlapping tiles. 36KYB (the clearer of the two) is used
    where it has data; 36KYA fills the rest after matching its mean and spread, band by band, to 36KYB on
    the pixels both cover, so there is no seam."""
    key = f"t9-{S2_DATE}-{left}-{bottom}-{right}-{top}"
    cached = S2_CACHE / f"{key}.npy"
    if cached.exists():
        return np.load(cached)
    import rasterio
    from rasterio.windows import from_bounds
    w, h = round((right - left) / 10), round((top - bottom) / 10)
    tiles = []
    with rasterio.Env(GDAL_DISABLE_READDIR_ON_OPEN="EMPTY_DIR", AWS_NO_SIGN_REQUEST="YES"):
        for sid in S2_SCENES:
            url = "/vsicurl/" + S2_URL.format(sq=sid.split("_")[1][3:], id=sid)
            with rasterio.open(url) as s:
                assert s.crs.to_epsg() == UTM
                win = from_bounds(left, bottom, right, top, transform=s.transform)
                tiles.append(s.read(window=win, boundless=True, fill_value=0, out_shape=(3, h, w)).astype(np.float32))
    b, a = tiles                      # a: 36KYB (primary), b: 36KYA (fill)
    va, vb = a.sum(0) > 0, b.sum(0) > 0
    both = va & vb
    out = a.copy()
    if both.any():
        for k in range(3):
            ma, sa, mb, sb = a[k][both].mean(), a[k][both].std(), b[k][both].mean(), b[k][both].std()
            b[k] = (b[k] - mb) * (sa / max(sb, 1e-6)) + ma
    fill = ~va & vb
    out[:, fill] = b[:, fill]
    out = np.clip(out, 0, 255).astype(np.uint8)
    S2_CACHE.mkdir(parents=True, exist_ok=True)
    np.save(cached, out)
    return out


def enhance(arr, gamma=0.9):
    """Gentle stretch for display, the same for all three bands so colours keep their balance."""
    lo, hi = np.percentile(arr, (0.5, 99.7))
    out = np.clip((arr.astype(np.float32) - lo) / max(hi - lo, 1), 0, 1) ** gamma
    return (out * 255).astype(np.uint8)


def route_view(line, segs, bounds, out_name, scale=1, mode="hero"):
    left, bottom, right, top = bounds
    raw = s2_window(*bounds)
    assert (raw.sum(0) == 0).mean() < 1e-3, f"{out_name}: the Sentinel-2 window has an area without data"
    base = Image.fromarray(np.transpose(enhance(raw, 1.0 if mode == "hero" else 0.9), (1, 2, 0)))
    if scale != 1:
        base = base.resize((base.width * scale, base.height * scale), Image.LANCZOS)
    S = 2
    W, H = base.size
    ppm = W / (right - left)
    ov = Image.new("RGBA", (W * S, H * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    px = lambda x, y: ((x - left) * ppm * S, (top - y) * ppm * S)
    if mode == "hero":
        band = [px(*xy) for xy in line.buffer(100, quad_segs=24).exterior.coords]
        d.polygon(band, fill=AMBER + (90,))
        route = [px(*xy) for xy in line.coords]
        d.line(route, fill=INK + (200,), width=12 * S, joint="curve")
        d.line(route, fill=ORANGE + (255,), width=6 * S, joint="curve")
    else:
        route = [px(*xy) for xy in line.coords]
        d.line(route, fill=INK + (210,), width=16 * S, joint="curve")
        for g in segs:
            part = substring(line, g["from_m"], g["to_m"])
            col = {"high": RED, "medium": AMBER, "low": LOWGREY}[g["rating"]]
            d.line([px(*xy) for xy in part.coords], fill=col + (255,), width=9 * S, joint="curve")
        f = font(19 * S, 750)
        for kk in range(0, int(line.length // 1000) + 1, 2):
            p = line.interpolate(kk * 1000)
            cx, cy = px(p.x, p.y)
            d.ellipse([cx - 7 * S, cy - 7 * S, cx + 7 * S, cy + 7 * S], fill=(255, 255, 255), outline=INK, width=3 * S)
            t = f"km {kk}"
            tw = d.textlength(t, font=f)
            ty = cy - 46 * S
            d.rounded_rectangle([cx - tw / 2 - 8 * S, ty - 4 * S, cx + tw / 2 + 8 * S, ty + 26 * S], radius=6 * S,
                                fill=INK + (215,))
            d.text((cx - tw / 2, ty - 1 * S), t, font=f, fill=(255, 255, 255))
        # scale bar, 1 km
        f2 = font(16 * S, 650)
        x0, y0 = 30 * S, (H - 40) * S
        d.rounded_rectangle([x0 - 12 * S, y0 - 34 * S, x0 + 1000 * ppm * S + 50 * S, y0 + 20 * S], radius=6 * S,
                            fill=INK + (200,))
        d.rectangle([x0, y0, x0 + 1000 * ppm * S, y0 + 7 * S], fill=(255, 255, 255))
        d.text((x0, y0 - 28 * S), "1 km", font=f2, fill=(255, 255, 255))
    ov = ov.resize((W, H), Image.LANCZOS)
    img = Image.alpha_composite(base.convert("RGBA"), ov).convert("RGB")
    OUT_IMG.mkdir(parents=True, exist_ok=True)
    img.save(OUT_IMG / out_name, "JPEG", quality=90, optimize=True, progressive=True)
    return list(img.size)


def main():
    line = load_route()
    marks = load_marks(line)
    segs = segments(line, marks)
    within = {b: sum(m["dist"] <= b for m in marks) for b in BUFFERS}
    minx, miny, maxx, maxy = line.bounds
    sizes = {}

    # Hero: 24 x 13.5 km from the western edge of the tiles, which puts the route in the right half so
    # the text on the left sits on open bush.
    x0 = 699960
    cy = round((miny + maxy) / 20) * 10
    assert x0 + 24000 > maxx + 1000
    sizes["t9-route-hero"] = route_view(line, segs, (x0, cy - 6750, x0 + 24000, cy + 6750), "t9-route-hero.jpg")
    # Overview: the route with 700 m around it, each 500 m coloured by its rating; 2x for sharp lines.
    ob = (round((minx - 700) / 10) * 10, round((miny - 1100) / 10) * 10,
          round((maxx + 700) / 10) * 10, round((maxy + 1100) / 10) * 10)
    sizes["t9-route-ratings"] = route_view(line, segs, ob, "t9-route-ratings.jpg", scale=2, mode="ratings")

    k0, k1 = EXCERPT_KM[0] * 1000, EXCERPT_KM[1] * 1000
    by_rating = {"high": 5500, "medium": 9000, "low": 7500}          # one full 500 m segment of each
    for lang, sfx in (("en", ""), ("pt", "-pt")):
        sizes[f"t9-register-km5-6{sfx}"], ids = strip(marks, segs, k0, k1, f"t9-register-km5-6{sfx}.jpg", lang,
                                                      width=1800, ppm_x=1640 / (k1 - k0))
        for rating, start in by_rating.items():
            g = next(s for s in segs if s["from_m"] == start)
            assert g["rating"] == rating
            sizes[f"t9-register-{rating}{sfx}"], _ = strip(marks, segs, 0, 0, f"t9-register-{rating}{sfx}.jpg",
                                                            lang, width=1000, ppm_x=1.4, labels=False,
                                                            badge=(rating, g["count"]), zone=(start, start + 500))

    shown = [m for m in marks if k0 <= m["chain"] <= k1 and m["dist"] <= ACROSS_M]
    data = {
        "route": "T-9 replacement pipeline",
        "province": "Inhambane, Mozambique",
        "route_length_km": round(line.length / 1000, 2),
        "buffers_m": list(BUFFERS),
        "method": "Reviewer-marked structures (manual review), no automatic detection in this sample",
        "imagery": ("Marks placed by a reviewer on a web basemap with no capture date; that basemap is not "
                    "reproduced. Route views use Copernicus Sentinel-2 L2A of " + S2_DATE + " (10 m), location only."),
        "sentinel2": {"date": S2_DATE, "scenes": S2_SCENES},
        "marks_total": len(marks),
        "within_m": {str(k): v for k, v in within.items()},
        "segment_m": SEGMENT,
        "rating_rule": {"high_over": HIGH_OVER, "buffer_m": max(BUFFERS)},
        "segments": segs,
        "segments_by_rating": {r: sum(s["rating"] == r for s in segs) for r in ("high", "medium", "low")},
        "excerpt_km": list(EXCERPT_KM),
        "register_excerpt": [{"id": m["id"], "km": round(m["chain"] / 1000, 2), "distance_m": round(m["dist"]),
                              "band": m["band"]} for m in shown],
        "images": sizes,
        "source_job": "corridor_909497cd (manual review, complete)",
    }
    OUT_DATA.parent.mkdir(parents=True, exist_ok=True)
    OUT_DATA.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({k: data[k] for k in ("marks_total", "within_m", "segments_by_rating")}, ensure_ascii=False))
    print("register excerpt:", len(shown), "marks,", shown[0]["id"], "to", shown[-1]["id"])


if __name__ == "__main__":
    main()
