#!/usr/bin/env python3
"""Rebuild the sample pipeline images and register from the app's stored job.

    /opt/favhousecheck/.venv/bin/python3 site/tools/make_samples.py [--tiles DIR] [--no-fetch]

The reviewer placed the sample's marks on Google satellite imagery in the app. What this tool draws:

  register views   strip maps drawn from the register alone: each reviewer mark by chainage and by
                   signed distance from the route (north side up), over the 50 m and 100 m bands.
                   No imagery and no coordinates. English and Portuguese (-pt) versions.
  route views      the route on a dated Copernicus Sentinel-2 L2A scene (10 m, free and open data,
                   credited "Contains modified Copernicus Sentinel data 2026"). 10 m pixels cannot
                   show individual structures, so no marks are drawn on it: the hero shows the route
                   and its 100 m band, the overview colours each 500 m of route by its rating.
  Google views     the reviewer's marks, the route and its 50 m and 100 m bands on Google satellite
                   imagery (owner decision 2026-09-27: allowed, with "Imagery © Google" on the image
                   and in the caption): five close-ups and one overview of the whole route, English
                   and Portuguese (-pt). No names, coordinates or chainage labels on the images.

  site/images/samples/*.jpg          masters for the image pipeline (build.py makes AVIF/WebP)
  site/images/samples/google/*.jpg   the Google-imagery masters (build.py refuses one that its data
                                     file does not declare, and any page that shows one without the
                                     attribution in the same figure)
  site/data/samples/sample-pipeline.json
                                route facts, 500 m segment ratings, the register excerpt and the text
                                drawn on each image (the build runs its guards on that text too)
  site/data/samples/sample-pipeline-google.json
                                the Google views: frames, counts, captions, alt text, attribution and
                                the text drawn on each image, in English and Portuguese

The site describes the sample only as "a high-pressure gas pipeline in Mozambique" (owner decision
2026-09-27): no route, field or province name goes into the images, the file names or the data.

Read-only against /opt/favhousecheck. The Sentinel-2 windows are read from the public
sentinel-cogs bucket (AWS Open Data) once and kept in site/.cache/s2/. Google tiles come from the
app's tile cache (/opt/favhousecheck/sat_cache/google, the tiles the review used) when it has them;
the rest are fetched once, one request at a time with a browser User-Agent and nothing else about
the requester, and kept in --tiles (default site/.cache/google-tiles). --no-fetch fails instead.
Output is deterministic for the same tiles: rerunning it leaves every file unchanged.
"""
import argparse
import io
import json
import math
import re
import time
import urllib.request
import zipfile
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFont
from pyproj import Transformer
from shapely.geometry import LineString, Point, box
from shapely.ops import substring
from shapely.ops import transform as shp_transform

# Every string drawn on an image is recorded, so the build can run its guards (withdrawn names,
# overstatement, pricing) on the images' own text as well as the page text.
_DRAWN = []
_draw_text = ImageDraw.ImageDraw.text


def _recording_text(self, xy, text, *args, **kwargs):
    _DRAWN.append(str(text))
    return _draw_text(self, xy, text, *args, **kwargs)


ImageDraw.ImageDraw.text = _recording_text


def take_drawn():
    out = list(dict.fromkeys(s for s in _DRAWN if s.strip()))
    _DRAWN.clear()
    return out


SITE = Path(__file__).resolve().parent.parent
JOB = Path("/opt/favhousecheck/results/corridor_909497cd")      # manual review job, 59 marks
ROUTE_DIR = Path("/opt/favhousecheck/uploads/corridor_909497cd")   # the job's uploaded route (one .kmz)
FONT = SITE / "fonts/inter-latin-wght-normal.woff2"
OUT_IMG = SITE / "images/samples"
OUT_DATA = SITE / "data/samples/sample-pipeline.json"
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
    "en": {"tag": "REGISTER VIEW · NO IMAGERY", "title": "Pipeline route · km {a}–{b}", "north": "north side",
           "south": "south side", "high": "High", "medium": "Medium", "low": "Low", "dec": "."},
    "pt": {"tag": "VISTA DO REGISTO · SEM IMAGENS", "title": "Traçado do gasoduto · km {a}–{b}", "north": "lado norte",
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
    (route,) = sorted(ROUTE_DIR.glob("*.kmz"))
    kml = zipfile.ZipFile(route).read("doc.kml").decode()
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
    key = f"sample-pipeline-{S2_DATE}-{left}-{bottom}-{right}-{top}"
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


# ------------------------------------------------------------------ Google satellite views
APP_TILES = Path("/opt/favhousecheck/sat_cache/google")                # {z}/{x}/{y}.jpg, read-only
TILE_URL = "https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}"        # favhousecheck/cache.py TILE_URLS["google"]
TILE_UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
TILE_GAP_S = 0.35              # at least this long between two tile requests
MERC = 20037508.342789244
G_OUT = OUT_IMG / "google"
G_DATA = SITE / "data/samples/sample-pipeline-google.json"
G_ZOOM, G_OVERVIEW_ZOOM = 20, 16
G_SIZE, G_MPP = (2400, 1600), 0.25          # close-ups: 600 m x 400 m on the ground, north up
G_OVERVIEW_SIZE = (2400, 820)
SAT_RED, SAT_AMBER, SAT_TEAL = (239, 68, 68), (251, 191, 36), (94, 234, 212)   # brighter on imagery

# Close-ups: (letter, chainage of the frame centre in m, shift east and north in m). Picked from the
# register, not from how the imagery looks: the stretches with most marks within 100 m (km 3.7 to 7.1),
# plus the two marks near km 9. Six views fill the two-column grid on the pages.
G_VIEWS = [("A", 4020, 0, 0), ("B", 4650, 0, 0), ("C", 5250, 0, 0), ("D", 5800, 0, 0), ("E", 6790, 0, 20),
           ("F", 9053, 0, 0)]

u2m = Transformer.from_crs(UTM, 3857, always_xy=True)
m2u = Transformer.from_crs(3857, UTM, always_xy=True)

GSTR = {
    "en": {
        "attribution": "Imagery © Google", "view": "View {k}", "route": "Pipeline route", "b50": "Within 50 m",
        "b100": "50 to 100 m", "beyond": "Beyond 100 m", "rings": "Rings: structures marked by a reviewer",
        "ring_key": "Reviewer mark", "dot_key": "Dots: reviewer marks", "band": "100 m band", "frames": "Close-up views",
        "badge": "Reviewed · manual marks", "dec": ".",
        "note": ("Each ring is centred on a mark the reviewer placed. Roofs without a ring were not marked in this "
                 "review: the views show the sample's marks exactly as recorded, with nothing added."),
        "summary": ("Together the close-ups show {shown} of the {total} reviewer-marked structures within 100 m of "
                    "the route."),
        "caption": "View {k} · km {a} to {b}",
        "credit": ("Imagery © Google, capture date not stated. Reviewer marks only, no automatic detections: "
                   "a person reviewed every result."),
        "ov_caption": "The whole route, with close-up views {first} to {last} outlined",
        "ov_credit": ("Imagery © Google, capture date not stated. Route, 100 m band and the reviewer's {n} marks "
                      "drawn by AfriScan; a person reviewed every result."),
    },
    "pt": {
        "attribution": "Imagens © Google", "view": "Vista {k}", "route": "Traçado do gasoduto", "b50": "Até 50 m",
        "b100": "50 a 100 m", "beyond": "Além de 100 m", "rings": "Círculos: construções marcadas pelo revisor",
        "ring_key": "Marcação do revisor", "dot_key": "Pontos: marcações do revisor", "band": "Faixa de 100 m", "frames": "Vistas de perto",
        "badge": "Revisto · marcação manual", "dec": ",",
        "note": ("Cada círculo está centrado numa marcação feita pelo revisor. As coberturas sem círculo não foram "
                 "marcadas nesta revisão: as vistas mostram as marcações do exemplo exactamente como foram "
                 "registadas, sem acrescentos."),
        "summary": ("No conjunto, as vistas de perto mostram {shown} das {total} construções marcadas pelo revisor a "
                    "menos de 100 m do traçado."),
        "caption": "Vista {k} · km {a} a {b}",
        "credit": ("Imagens © Google, sem data de captação indicada. Só marcações do revisor, sem detecções "
                   "automáticas: uma pessoa reviu todos os resultados."),
        "ov_caption": "Todo o traçado, com as vistas de perto {first} a {last} assinaladas",
        "ov_credit": ("Imagens © Google, sem data de captação indicada. Traçado, faixa de 100 m e as {n} marcações "
                      "do revisor desenhados pela AfriScan; uma pessoa reviu todos os resultados."),
    },
}


def g_count_text(lang, n50, n100, nb):
    """The figure text under each close-up, from the counts in that view."""
    if lang == "en":
        s = lambda n, one, many: f"{n} {one if n == 1 else many}"
        if n100 == 0:
            t = "No reviewer-marked structures within 100 m of the route in this view"
        else:
            t = f"{s(n100, 'reviewer-marked structure', 'reviewer-marked structures')} within 100 m of the route in this view"
            t += (", none of them within 50 m" if n50 == 0 else ", within 50 m" if n50 == n100 == 1
                  else ", all of them within 50 m" if n50 == n100 else f", {n50} of them within 50 m")
        if nb:
            t += f"; {s(nb, 'more', 'more')} beyond 100 m"
        return t + "."
    if n100 == 0:
        t = "Nenhuma construção marcada pelo revisor a menos de 100 m do traçado nesta vista"
    else:
        noun = "construção marcada" if n100 == 1 else "construções marcadas"
        t = f"{n100} {noun} pelo revisor a menos de 100 m do traçado nesta vista"
        t += (", nenhuma a menos de 50 m" if n50 == 0 else ", a menos de 50 m" if n50 == n100 == 1
              else ", todas a menos de 50 m" if n50 == n100 else f", {'uma' if n50 == 1 else n50} delas a menos de 50 m")
    if nb:
        t += f"; mais {nb} além dos 100 m"
    return t + "."


G_ALT = {
    "en": ("Satellite view of about 600 by 400 m along the pipeline route, north up: the route as an orange line, the "
           "50 m band shaded red and the 100 m band shaded amber on both sides, and {n} rings on the structures the "
           "reviewer marked ({c50} within 50 m, {c100} between 50 and 100 m, {cb} beyond 100 m). {scene}"),
    "pt": ("Vista de satélite de cerca de 600 por 400 m ao longo do traçado do gasoduto, com o norte em cima: o traçado "
           "como uma linha cor de laranja, a faixa de 50 m sombreada a vermelho e a de 100 m a âmbar dos dois lados, e "
           "{n} círculos nas construções que o revisor marcou ({c50} a menos de 50 m, {c100} entre 50 e 100 m, {cb} "
           "além de 100 m). {scene}"),
}
# What each close-up shows, written from the images (read them after any change to G_VIEWS).
G_SCENE = {
    "A": {"en": ("The route follows a dirt track west to east through small fields and scattered homesteads, with a "
                 "larger cluster of homesteads to the north."),
          "pt": ("O traçado segue uma picada de oeste para leste entre pequenas machambas e habitações dispersas, com "
                 "um aglomerado maior de habitações a norte.")},
    "B": {"en": ("The route follows the same track east and bends where it meets a paved road near the eastern edge; "
                 "homesteads, fields and trees line both sides."),
          "pt": ("O traçado continua pela mesma picada e faz uma curva junto de uma estrada asfaltada, perto do limite "
                 "leste; há habitações, machambas e árvores dos dois lados.")},
    "C": {"en": ("The route bends from east to north-east beside a dirt road, with homesteads, fields and palm trees "
                 "on both sides of the bend."),
          "pt": ("O traçado faz uma curva de leste para nordeste junto de uma estrada de terra, com habitações, "
                 "machambas e palmeiras dos dois lados da curva.")},
    "D": {"en": ("The route runs north-east beside a wide dirt track, with homesteads and palm trees to the west and "
                 "bush and woodland to the south-east."),
          "pt": ("O traçado segue para nordeste ao lado de uma picada larga, com habitações e palmeiras a oeste e mato "
                 "e floresta a sudeste.")},
    "E": {"en": ("The route runs north-east beside the track through fields and bush, with a cluster of homesteads "
                 "to the west."),
          "pt": ("O traçado segue para nordeste ao lado da picada, entre machambas e mato, com um aglomerado de "
                 "habitações a oeste.")},
    "F": {"en": "The route runs west to east along the track through bush and grassland, with a wetland to the north.",
          "pt": "O traçado segue de oeste para leste pela picada, entre mato e capim, com uma zona húmida a norte."},
}
G_OV_ALT = {
    "en": ("Satellite overview of the whole pipeline route, about 10 km from west to east, north up: the route as an "
           "orange line with its 100 m band, the reviewer's {n} marks as dots coloured by distance band, and {k} "
           "outlined frames, {first} to {last}, for the close-up views. {scene}"),
    "pt": ("Vista geral de satélite de todo o traçado do gasoduto, com cerca de 10 km de oeste para leste e o norte em "
           "cima: o traçado como uma linha cor de laranja com a faixa de 100 m, as {n} marcações do revisor como "
           "pontos coloridos pela faixa de distância, e {k} molduras, de {first} a {last}, para as vistas de perto. "
           "{scene}"),
}
G_OV_SCENE = {"en": "It starts at a gas facility in the west, passes fields and settlements, and ends beside a wetland in the east.",
              "pt": ("Começa numa instalação de gás a oeste, passa por machambas e povoações e termina junto de uma "
                     "zona húmida a leste.")}


class Tiles:
    """Google XYZ tiles: the app's cache first (read-only), then this tool's cache, then one polite fetch."""

    def __init__(self, cache_dir, fetch=True):
        self.dir, self.fetch, self.fetched, self.last = Path(cache_dir), fetch, 0, 0.0

    def get(self, z, x, y):
        for p in (APP_TILES / f"{z}/{x}/{y}.jpg", self.dir / f"{z}/{x}/{y}.jpg"):
            if p.exists():
                return Image.open(p).convert("RGB")
        if not self.fetch:
            raise SystemExit(f"tile {z}/{x}/{y} is in no cache and --no-fetch is set")
        wait = TILE_GAP_S - (time.monotonic() - self.last)
        if wait > 0:
            time.sleep(wait)
        req = urllib.request.Request(TILE_URL.format(x=x, y=y, z=z),
                                     headers={"User-Agent": TILE_UA, "Accept": "image/avif,image/webp,image/*,*/*;q=0.8"})
        for attempt in range(3):
            try:
                with urllib.request.urlopen(req, timeout=30) as r:
                    body, kind = r.read(), r.headers.get("Content-Type", "")
                break
            except OSError:
                if attempt == 2:
                    raise
                time.sleep(5 * (attempt + 1))
        self.last = time.monotonic()
        self.fetched += 1
        if not kind.startswith("image/"):
            raise SystemExit(f"tile {z}/{x}/{y}: got {kind!r}, not an image")
        im = Image.open(io.BytesIO(body)).convert("RGB")
        out = self.dir / f"{z}/{x}/{y}.jpg"
        out.parent.mkdir(parents=True, exist_ok=True)
        out.with_suffix(".tmp").write_bytes(body)
        out.with_suffix(".tmp").rename(out)
        return im


def g_mosaic(tiles, z, bounds, size):
    """Web Mercator window (left, bottom, right, top in EPSG:3857 metres) resampled to SIZE pixels."""
    left, bottom, right, top = bounds
    s = 256 * 2 ** z / (2 * MERC)
    px0, px1, py0, py1 = (left + MERC) * s, (right + MERC) * s, (MERC - top) * s, (MERC - bottom) * s
    tx0, tx1, ty0, ty1 = int(px0 // 256), int((px1 - 1e-6) // 256), int(py0 // 256), int((py1 - 1e-6) // 256)
    canvas = Image.new("RGB", ((tx1 - tx0 + 1) * 256, (ty1 - ty0 + 1) * 256))
    for x in range(tx0, tx1 + 1):
        for y in range(ty0, ty1 + 1):
            canvas.paste(tiles.get(z, x, y), ((x - tx0) * 256, (y - ty0) * 256))
    return canvas.resize(size, Image.LANCZOS, box=(px0 - tx0 * 256, py0 - ty0 * 256, px1 - tx0 * 256, py1 - ty0 * 256))


def merc_lat(y):
    """Latitude in degrees of a Web Mercator northing."""
    return math.degrees(math.atan(math.sinh(y / 6378137.0)))


def g_frame(line, chain, dx, dy, size, mpp):
    c = line.interpolate(chain)
    cx, cy = u2m.transform(c.x + dx, c.y + dy)
    k = 1 / math.cos(math.radians(Transformer.from_crs(UTM, 4326, always_xy=True).transform(c.x, c.y)[1]))
    hw, hh = size[0] * mpp * k / 2, size[1] * mpp * k / 2
    return (cx - hw, cy - hh, cx + hw, cy + hh)


def g_polys(geom):
    return [g for g in getattr(geom, "geoms", [geom]) if not g.is_empty]


def band_of(d):
    return "a" if d <= 50 else "b" if d <= 100 else "c"


def g_hits(rect, circles, margin=0):
    """The circles (x, y, r) that overlap rect (x0, y0, x1, y1), all in drawing pixels."""
    x0, y0, x1, y1 = rect
    return [(x, y, r) for x, y, r in circles
            if x0 - r - margin < x < x1 + r + margin and y0 - r - margin < y < y1 + r + margin]


def g_panel(d, S, x0, y0, rows, title=None, fs=30, draw=True, anchor="left"):
    """Legend panel; returns its box. rows: (symbol, label); symbols: route, a, b, c (band swatch with ring),
    ring (a plain ring), dot-a/b/c, frame, note. Everything scales with the font size fs (image pixels). anchor="right": x0 is the
    panel's right edge. draw=False only measures."""
    u = fs / 30
    q = lambda v: round(v * u) * S
    f, ft = font(fs * S, 600), font(round(fs * 1.2) * S, 750)
    pad, rh, sw, th = q(22), round(fs * 1.5) * S, round(fs * 1.9) * S, round(fs * 1.75) * S
    w = max([d.textlength(t, font=f) + (0 if s == "note" else sw + q(16)) for s, t in rows]
            + ([d.textlength(title, font=ft)] if title else []))
    h = len(rows) * rh + (th if title else 0)
    if anchor == "right":
        x0 -= w + 2 * pad
    box_ = (x0, y0, x0 + w + 2 * pad, y0 + h + 2 * pad)
    if not draw:
        return box_
    d.rounded_rectangle(box_, radius=q(14), fill=INK + (205,))
    y = y0 + pad
    if title:
        d.text((x0 + pad, y + th / 2 - q(2)), title, font=ft, fill=(255, 255, 255), anchor="lm")
        y += th
    for sym, label in rows:
        cy, sx = y + rh / 2, x0 + pad
        col = {"a": SAT_RED, "b": SAT_AMBER, "c": SAT_TEAL}.get(sym[-1])
        if sym == "route":
            d.line([sx, cy, sx + sw, cy], fill=INK + (255,), width=q(12))
            d.line([sx, cy, sx + sw, cy], fill=ORANGE + (255,), width=q(7))
        elif sym in ("a", "b"):
            d.rectangle([sx, cy - rh * 0.32, sx + sw, cy + rh * 0.32], fill=col + (110,), outline=col + (255,), width=q(2))
        if sym in ("a", "b", "c"):
            r = rh * 0.3
            d.ellipse([sx + sw / 2 - r, cy - r, sx + sw / 2 + r, cy + r], outline=col + (255,), width=q(5))
        elif sym == "ring":
            r = rh * 0.3
            d.ellipse([sx + sw / 2 - r, cy - r, sx + sw / 2 + r, cy + r], outline=INK + (255,), width=q(9))
            d.ellipse([sx + sw / 2 - r + q(2), cy - r + q(2), sx + sw / 2 + r - q(2), cy + r - q(2)],
                      outline=(255, 255, 255, 255), width=q(4))
        elif sym.startswith("dot"):
            r = rh * 0.2
            d.ellipse([sx + sw / 2 - r, cy - r, sx + sw / 2 + r, cy + r], fill=col + (255,), outline=INK + (255,),
                      width=q(2))
        elif sym == "band":
            d.rectangle([sx, cy - rh * 0.3, sx + sw, cy + rh * 0.3], fill=SAT_AMBER + (120,), outline=SAT_AMBER + (255,),
                        width=q(2))
        elif sym == "frame":
            d.rectangle([sx + q(4), cy - rh * 0.3, sx + sw - q(4), cy + rh * 0.3], outline=(255, 255, 255, 255),
                        width=q(4))
        tx = sx if sym == "note" else sx + sw + q(16)
        d.text((tx, cy), label, font=f, fill=(255, 255, 255) if sym != "note" else (214, 220, 228), anchor="lm")
        y += rh
    return box_


def g_furniture(d, S, W, H, lang, scale_m, ppm, fs=30, scale_right=False, att_px=None):
    """North arrow and scale bar (bottom left, or left of the attribution) and the attribution (bottom
    right), on dark pills. The attribution is set a size larger than the rest (att_px image pixels,
    default 1.25 fs), so it stays legible where the image is shown small."""
    L = GSTR[lang]
    u = fs / 30
    q = lambda v: round(v * u) * S
    att_px = att_px or round(fs * 1.25)
    f, fa = font(fs * S, 650), font(att_px * S, 700)
    x1, y1 = W * S - q(22), H * S - q(22)
    # attribution
    t = L["attribution"]
    ph = round(att_px * 1.64) * S
    tw = d.textlength(t, font=fa)
    boxes = [(x1 - tw - q(36), y1 - ph, x1, y1)]
    d.rounded_rectangle(boxes[0], radius=q(10), fill=(0, 0, 0, 185))
    d.text((x1 - q(18), y1 - ph / 2), t, font=fa, fill=(255, 255, 255), anchor="rm")
    # north arrow and scale bar
    bar = scale_m * ppm * S
    label = f"{scale_m} m" if scale_m < 1000 else f"{scale_m // 1000} km"
    bh = round(fs * 3.2) * S
    pill_w = q(38) + q(44) + max(bar, d.textlength(label, font=f)) + q(24)
    x0 = (x1 - tw - q(36) - q(16) - pill_w) if scale_right else q(22)
    top = y1 - bh
    ax = x0 + q(38)
    bx = ax + q(44)
    boxes.append((x0, top, x0 + pill_w, y1))
    d.rounded_rectangle(boxes[1], radius=q(10), fill=(0, 0, 0, 175))
    d.text((ax, top + q(22)), "N", font=font(round(fs * 0.85) * S, 750), fill=(255, 255, 255), anchor="mm")
    tip, base = top + q(40), y1 - q(14)
    d.polygon([(ax, tip), (ax + q(14), base), (ax, base - q(11)), (ax - q(14), base)], fill=(255, 255, 255))
    by = y1 - q(18)
    d.rectangle([bx, by - q(9), bx + bar, by], fill=(255, 255, 255))
    for tx in (bx, bx + bar):
        d.rectangle([tx - q(2), by - q(22), tx + q(2), by], fill=(255, 255, 255))
    d.text((bx, by - q(32)), label, font=f, fill=(255, 255, 255), anchor="ls")
    return boxes


def g_save(img, name):
    """JPEG with no EXIF, XMP, ICC or comment: a fresh image built from the pixels alone."""
    clean = Image.frombytes("RGB", img.size, img.convert("RGB").tobytes())
    G_OUT.mkdir(parents=True, exist_ok=True)
    path = G_OUT / f"{name}.jpg"
    clean.save(path, "JPEG", quality=90, optimize=True, progressive=True)
    with Image.open(path) as chk:
        extra = set(chk.info) - {"jfif", "jfif_version", "jfif_unit", "jfif_density", "progressive", "progression"}
        assert not extra and not chk.getexif(), f"{path}: metadata left in the file: {sorted(extra)}"
    return path


def g_view(tiles, line, marks, letter, bounds, lang, size=G_SIZE):
    """One close-up: imagery, 100 m and 50 m bands, the route, a ring on each reviewer mark, legend."""
    W, H = size
    L = GSTR[lang]
    img = g_mosaic(tiles, G_ZOOM, bounds, size)
    S = 2
    left, bottom, right, top = bounds
    px = lambda X, Y: ((X - left) / (right - left) * W * S, (top - Y) / (top - bottom) * H * S)
    to_m = lambda g: shp_transform(lambda x, y, z=None: u2m.transform(x, y), g)
    area = shp_transform(lambda x, y, z=None: m2u.transform(x, y), box(*bounds))
    near = line.intersection(area.buffer(300))            # the route within 300 m of the frame: no end caps in view
    b50, b100 = near.buffer(50, quad_segs=32), near.buffer(100, quad_segs=32)
    rings = lambda geom: [[px(*xy) for xy in to_m(p).exterior.coords] for p in g_polys(geom)]
    holes = lambda geom: [[px(*xy) for xy in to_m(i).coords] for p in g_polys(geom) for i in p.interiors]
    masks = {}
    for k, geom in (("100", b100), ("50", b50)):
        m = Image.new("L", (W * S, H * S), 0)
        md = ImageDraw.Draw(m)
        for r in rings(geom):
            md.polygon(r, fill=255)
        for r in holes(geom):
            md.polygon(r, fill=0)
        masks[k] = m
    ov = Image.new("RGBA", (W * S, H * S), (0, 0, 0, 0))
    for col, mask, alpha in ((SAT_AMBER, ImageChops.subtract(masks["100"], masks["50"]), 44), (SAT_RED, masks["50"], 50)):
        layer = Image.new("RGBA", ov.size, col + (0,))
        layer.putalpha(mask.point(lambda v, a=alpha: a if v else 0))
        ov = Image.alpha_composite(ov, layer)
    d = ImageDraw.Draw(ov)
    for geom, col in ((b100, SAT_AMBER), (b50, SAT_RED)):
        for r in rings(geom) + holes(geom):
            d.line(r + [r[0]], fill=INK + (120,), width=6 * S)
            d.line(r + [r[0]], fill=col + (255,), width=3 * S)
    for part in g_polys(near):
        pts = [px(*u2m.transform(*xy)) for xy in part.coords]
        d.line(pts, fill=INK + (215,), width=14 * S, joint="curve")
        d.line(pts, fill=ORANGE + (255,), width=7 * S, joint="curve")
    shown, circles = [], []
    R = 36 * S
    for m in marks:
        x, y = px(*u2m.transform(m["p"].x, m["p"].y))
        if 0 <= x <= W * S and 0 <= y <= H * S:
            shown.append(m)
            circles.append((x, y, R))
            col = {"a": SAT_RED, "b": SAT_AMBER, "c": SAT_TEAL}[band_of(m["dist"])]
            d.ellipse([x - R, y - R, x + R, y + R], outline=INK + (215,), width=14 * S)
            d.ellipse([x - R + 3 * S, y - R + 3 * S, x + R - 3 * S, y + R - 3 * S], outline=col + (255,), width=7 * S)
    # The legend goes top left, or top right when a ring would sit under it; no ring may hide under the
    # legend, the scale bar or the attribution (move the frame in G_VIEWS if one does).
    # Measured in every language, so the English and Portuguese views put it in the same corner.
    legend = lambda T: ([("ring", T["ring_key"]), ("route", T["route"]), ("a", T["b50"]), ("b", T["b100"]),
                         ("c", T["beyond"])], T["view"].format(k=letter))
    rows, title = legend(L)
    spots = [(24 * S, "left"), (W * S - 24 * S, "right")]
    free = [sp for sp in spots
            if not any(g_hits(g_panel(d, S, sp[0], 24 * S, *legend(T), 44, draw=False, anchor=sp[1]), circles, 8 * S)
                       for T in GSTR.values())]
    if not free:
        raise SystemExit(f"view {letter}: a reviewer mark sits under the legend in both top corners; move the frame")
    g_panel(d, S, free[0][0], 24 * S, rows, title=title, fs=44, anchor=free[0][1])
    # The scale bar sits bottom left, or beside the attribution when a ring would sit under it.
    ppm = W / ((right - left) * math.cos(math.radians(merc_lat((top + bottom) / 2))))
    for scale_right in (False, True):
        layer = Image.new("RGBA", ov.size, (0, 0, 0, 0))
        if not any(g_hits(b_, circles, 8 * S) for b_ in g_furniture(ImageDraw.Draw(layer), S, W, H, lang, 100, ppm,
                                                                     fs=44, scale_right=scale_right)):
            break
    else:
        raise SystemExit(f"view {letter}: a reviewer mark sits under the scale bar or the attribution; move the frame")
    ov = Image.alpha_composite(ov, layer)
    ov = ov.resize((W, H), Image.LANCZOS)
    out = Image.alpha_composite(img.convert("RGBA"), ov)
    # chainage covered by the frame: the route's own points inside it, every 10 m
    inside = [c for c in range(0, int(line.length) + 1, 10) if area.contains(line.interpolate(c))]
    return out, shown, (min(inside), max(inside))


def g_overview(tiles, line, marks, frames, lang):
    W, H = G_OVERVIEW_SIZE
    L = GSTR[lang]
    minx, miny, maxx, maxy = line.bounds
    c = Point((minx + maxx) / 2, (miny + maxy) / 2)
    cx, cy = u2m.transform(c.x, c.y)
    k = 1 / math.cos(math.radians(merc_lat(cy)))
    ground_w = (maxx - minx) + 1200
    hw, hh = ground_w * k / 2, ground_w * H / W * k / 2
    bounds = (cx - hw, cy - hh, cx + hw, cy + hh)
    assert (maxy - miny) + 600 < 2 * hh / k, "overview: the route does not fit the frame height"
    img = g_mosaic(tiles, G_OVERVIEW_ZOOM, bounds, (W, H))
    S = 2
    left, bottom, right, top = bounds
    px = lambda X, Y: ((X - left) / (right - left) * W * S, (top - Y) / (top - bottom) * H * S)
    to_m = lambda g: shp_transform(lambda x, y, z=None: u2m.transform(x, y), g)
    ov = Image.new("RGBA", (W * S, H * S), (0, 0, 0, 0))
    m = Image.new("L", ov.size, 0)
    md = ImageDraw.Draw(m)
    for p in g_polys(line.buffer(100, quad_segs=24)):
        md.polygon([px(*xy) for xy in to_m(p).exterior.coords], fill=255)
    layer = Image.new("RGBA", ov.size, SAT_AMBER + (0,))
    layer.putalpha(m.point(lambda v: 80 if v else 0))
    ov = Image.alpha_composite(ov, layer)
    d = ImageDraw.Draw(ov)
    pts = [px(*u2m.transform(*xy)) for xy in line.coords]
    d.line(pts, fill=INK + (220,), width=10 * S, joint="curve")
    d.line(pts, fill=ORANGE + (255,), width=5 * S, joint="curve")
    circles = []
    for mk in sorted(marks, key=lambda q: -q["dist"]):
        x, y = px(*u2m.transform(mk["p"].x, mk["p"].y))
        col = {"a": SAT_RED, "b": SAT_AMBER, "c": SAT_TEAL}[band_of(mk["dist"])]
        r = 7 * S
        circles.append((x, y, r))
        d.ellipse([x - r, y - r, x + r, y + r], fill=col + (255,), outline=INK + (255,), width=2 * S)
    fl = font(34 * S, 800)
    for letter, (l_, b_, r_, t_) in frames:
        (x0, y0), (x1, y1) = px(l_, t_), px(r_, b_)
        d.rectangle([x0, y0, x1, y1], outline=INK + (200,), width=7 * S)
        d.rectangle([x0, y0, x1, y1], outline=(255, 255, 255, 255), width=3 * S)
        tw = d.textlength(letter, font=fl)
        bx, by = x0, y0 - 52 * S
        d.rounded_rectangle([bx, by, bx + tw + 24 * S, by + 48 * S], radius=8 * S, fill=INK + (225,))
        circles += [((x0 + x1) / 2, (y0 + y1) / 2, max(x1 - x0, y1 - y0) / 2), (bx + tw / 2 + 12 * S, by + 24 * S, 30 * S)]
        d.text((bx + 12 * S + tw / 2, by + 24 * S), letter, font=fl, fill=(255, 255, 255), anchor="mm")
    boxes = [g_panel(d, S, 22 * S, 22 * S, [("route", L["route"]), ("band", L["band"]), ("dot-a", L["b50"]),
                                            ("dot-b", L["b100"]), ("dot-c", L["beyond"]), ("frame", L["frames"]),
                                            ("note", L["dot_key"])], fs=28)]
    # The attribution matches the close-ups' (55 px): the overview is shown at the same width on phones.
    boxes += g_furniture(d, S, W, H, lang, 1000, W / ((right - left) / k), fs=32, scale_right=True, att_px=55)
    for b_ in boxes:
        if g_hits(b_, circles, 6 * S):
            raise SystemExit("overview: a mark or a close-up frame sits under the legend, the scale bar or the attribution")
    ov = ov.resize((W, H), Image.LANCZOS)
    return Image.alpha_composite(img.convert("RGBA"), ov), bounds


def google_views(line, marks, tiles):
    assert [v[0] for v in G_VIEWS] == list(G_SCENE), "write a G_SCENE entry for every view, from its image"
    km = lambda m, lang: f"{m / 1000:.1f}".replace(".", GSTR[lang]["dec"])
    data = {
        "sample": "sample-pipeline",
        "description": "Reviewer marks from the pipeline sample on Google satellite imagery",
        "imagery": {"provider": "google", "capture_date": None,
                    "attribution": {lang: GSTR[lang]["attribution"] for lang in GSTR},
                    "zoom": {"close-ups": G_ZOOM, "overview": G_OVERVIEW_ZOOM}},
        "marks": "reviewer marks only (manual review); no automatic detections",
        "marks_total": len(marks),
        "strings": {lang: {k: GSTR[lang][k] for k in ("attribution", "route", "b50", "b100", "beyond", "badge",
                                                      "rings", "note")} for lang in GSTR},
        "shown_within_100": None,
        "within_100_total": sum(m["dist"] <= 100 for m in marks),
        "views": [],
    }
    frames = []
    for letter, chain, dx, dy in G_VIEWS:
        bounds = g_frame(line, chain, dx, dy, G_SIZE, G_MPP)
        frames.append((letter, bounds))
        view = {"id": letter, "src": {}, "alt": {}, "caption": {}, "text": {}, "credit": {}, "drawn_text": {}}
        for lang, sfx in (("en", ""), ("pt", "-pt")):
            take_drawn()
            img, shown, (a, b) = g_view(tiles, line, marks, letter, bounds, lang)
            name = f"pipeline-view-{letter.lower()}{sfx}"
            g_save(img, name)
            n = {band: sum(band_of(m["dist"]) == band for m in shown) for band in "abc"}
            n50, n100, nb = n["a"], n["a"] + n["b"], n["c"]
            view["src"][lang] = f"samples/google/{name}"
            view["drawn_text"][lang] = take_drawn()
            view["caption"][lang] = GSTR[lang]["caption"].format(k=letter, a=km(a, lang), b=km(b, lang))
            view["text"][lang] = g_count_text(lang, n50, n100, nb)
            view["credit"][lang] = GSTR[lang]["credit"]
            view["alt"][lang] = G_ALT[lang].format(n=len(shown), c50=n50, c100=n["b"], cb=nb,
                                                    scene=G_SCENE[letter][lang]).strip()
        view.update(km=[round(a / 1000, 2), round(b / 1000, 2)], ground_m=[round(G_SIZE[0] * G_MPP), round(G_SIZE[1] * G_MPP)],
                    marks_in_view=len(shown), within_50=n50, within_100=n100, beyond_100=nb,
                    register_ids=[m["id"] for m in shown], within_100_ids=[m["id"] for m in shown if m["dist"] <= 100],
                    size=list(G_SIZE))
        data["views"].append(view)
    # The gallery fills in {shown} and {total} for the views a page shows (lib/sample_gallery.py).
    ids100 = {i for v in data["views"] for i in v["within_100_ids"]}
    data["shown_within_100"] = len(ids100)
    for lang in GSTR:
        data["strings"][lang]["summary"] = GSTR[lang]["summary"]
    first, last = G_VIEWS[0][0], G_VIEWS[-1][0]
    ov = {"id": "overview", "src": {}, "alt": {}, "caption": {}, "credit": {}, "drawn_text": {}}
    for lang, sfx in (("en", ""), ("pt", "-pt")):
        take_drawn()
        img, _ = g_overview(tiles, line, marks, frames, lang)
        name = f"pipeline-overview{sfx}"
        g_save(img, name)
        ov["src"][lang] = f"samples/google/{name}"
        ov["drawn_text"][lang] = take_drawn()
        ov["caption"][lang] = GSTR[lang]["ov_caption"].format(first=first, last=last)
        ov["credit"][lang] = GSTR[lang]["ov_credit"].format(n=len(marks))
        ov["alt"][lang] = G_OV_ALT[lang].format(n=len(marks), k=len(G_VIEWS), first=first, last=last,
                                                 scene=G_OV_SCENE[lang]).strip()
    ov.update(size=list(G_OVERVIEW_SIZE), marks_in_view=len(marks))
    data["overview"] = ov
    G_DATA.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    for v in data["views"]:
        print(f"view {v['id']}: km {v['km'][0]}-{v['km'][1]}, {v['marks_in_view']} marks in view, "
              f"{v['within_50']} within 50 m, {v['within_100']} within 100 m, {v['beyond_100']} beyond")
    print(f"close-ups show {data['shown_within_100']} of the {data['within_100_total']} marks within 100 m")
    print(f"tiles fetched: {tiles.fetched}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tiles", default=str(SITE / ".cache/google-tiles"),
                    help="where fetched Google tiles are kept (default site/.cache/google-tiles)")
    ap.add_argument("--no-fetch", action="store_true", help="use cached tiles only; fail if one is missing")
    args = ap.parse_args()
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
    texts = {}
    take_drawn()
    sizes["sample-pipeline-route-hero"] = route_view(line, segs, (x0, cy - 6750, x0 + 24000, cy + 6750),
                                                    "sample-pipeline-route-hero.jpg")
    texts["samples/sample-pipeline-route-hero"] = take_drawn()
    # Overview: the route with 700 m around it, each 500 m coloured by its rating; 2x for sharp lines.
    ob = (round((minx - 700) / 10) * 10, round((miny - 1100) / 10) * 10,
          round((maxx + 700) / 10) * 10, round((maxy + 1100) / 10) * 10)
    sizes["sample-pipeline-route-ratings"] = route_view(line, segs, ob, "sample-pipeline-route-ratings.jpg", scale=2,
                                                       mode="ratings")
    texts["samples/sample-pipeline-route-ratings"] = take_drawn()

    k0, k1 = EXCERPT_KM[0] * 1000, EXCERPT_KM[1] * 1000
    by_rating = {"high": 5500, "medium": 9000, "low": 7500}          # one full 500 m segment of each
    for lang, sfx in (("en", ""), ("pt", "-pt")):
        sizes[f"sample-pipeline-register-km5-6{sfx}"], ids = strip(
            marks, segs, k0, k1, f"sample-pipeline-register-km5-6{sfx}.jpg", lang, width=1800, ppm_x=1640 / (k1 - k0))
        texts[f"samples/sample-pipeline-register-km5-6{sfx}"] = take_drawn()
        for rating, start in by_rating.items():
            g = next(s for s in segs if s["from_m"] == start)
            assert g["rating"] == rating
            sizes[f"sample-pipeline-register-{rating}{sfx}"], _ = strip(
                marks, segs, 0, 0, f"sample-pipeline-register-{rating}{sfx}.jpg", lang, width=1000, ppm_x=1.4,
                labels=False, badge=(rating, g["count"]), zone=(start, start + 500))
            texts[f"samples/sample-pipeline-register-{rating}{sfx}"] = take_drawn()

    shown = [m for m in marks if k0 <= m["chain"] <= k1 and m["dist"] <= ACROSS_M]
    data = {
        "route": "High-pressure gas pipeline",
        "country": "Mozambique",
        "route_length_km": round(line.length / 1000, 2),
        "buffers_m": list(BUFFERS),
        "method": "Reviewer-marked structures (manual review), no automatic detection in this sample",
        "imagery": ("Marks placed by a reviewer on Google satellite imagery, which states no capture date; the "
                    "Google views (sample-pipeline-google.json) show them on it. Route views use Copernicus "
                    "Sentinel-2 L2A of " + S2_DATE + " (10 m), location only."),
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
        "image_text": texts,
        "source_job": "corridor_909497cd (manual review, complete)",
    }
    OUT_DATA.parent.mkdir(parents=True, exist_ok=True)
    OUT_DATA.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({k: data[k] for k in ("marks_total", "within_m", "segments_by_rating")}, ensure_ascii=False))
    print("register excerpt:", len(shown), "marks,", shown[0]["id"], "to", shown[-1]["id"])
    google_views(line, marks, Tiles(args.tiles, fetch=not args.no_fetch))


if __name__ == "__main__":
    main()
