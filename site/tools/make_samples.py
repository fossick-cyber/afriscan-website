#!/usr/bin/env python3
"""Rebuild the T-9 sample images and register from the app's stored job, in natural colour.

The images once published on afri-scan.com were the app's old annotated JPEGs, which wrote
rasterio RGB arrays through OpenCV as BGR, so red and blue were swapped. This tool starts from
the raw GeoTIFF chunks (checked: ColorInterp red/green/blue, natural colour), draws the route,
the 50 m and 100 m buffers and the reviewer's marks itself, and writes:

  site/images/samples/*.jpg        masters for the image pipeline (build.py makes AVIF/WebP)
  site/data/samples/t9.json        route facts, 500 m segment ratings and the register excerpt

Read-only against /opt/favhousecheck. Run it by hand when the sample should change:
  /opt/favhousecheck/.venv/bin/python3 site/tools/make_samples.py
"""
import json
import re
import zipfile
from pathlib import Path

import numpy as np
import rasterio
from PIL import Image, ImageDraw, ImageFont
from pyproj import Transformer
from rasterio.merge import merge
from shapely.geometry import LineString, Point, Polygon, MultiPolygon
from shapely.ops import substring

SITE = Path(__file__).resolve().parent.parent
JOB = Path("/opt/favhousecheck/results/corridor_909497cd")      # manual review job, 59 marks
ROUTE = Path("/opt/favhousecheck/uploads/corridor_909497cd/T-9 Replacement Pipeline Rev 00.kmz")
FONT = SITE / "fonts/inter-latin-wght-normal.woff2"
OUT_IMG = SITE / "images/samples"
OUT_DATA = SITE / "data/samples/t9.json"

BUFFERS = (50, 100)            # the job's buffers; the largest drives the segment rating
SEGMENT = 500
HIGH_OVER = 5                  # zones.py: high = more than 5 inside the largest buffer
UTM = 32736                    # the job's UTM zone (EPSG:32736, 36S)

C_ROUTE = (232, 103, 47)       # --accent
C_IN50 = (239, 68, 68)
C_IN100 = (245, 158, 11)
C_BEYOND = (94, 234, 212)      # --teal

to_utm = Transformer.from_crs(4326, UTM, always_xy=True)
to_ll = Transformer.from_crs(UTM, 4326, always_xy=True)


def font(size, weight=700):
    f = ImageFont.truetype(str(FONT), size)
    f.set_variation_by_axes([weight])
    return f


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
        lon, lat = f["geometry"]["coordinates"]
        p = Point(*to_utm.transform(lon, lat))
        marks.append({"lon": lon, "lat": lat, "p": p, "chain": line.project(p), "dist": line.distance(p)})
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


def mosaic(chunk_ids):
    srcs = [rasterio.open(JOB / f"sat_chunks/sat_chunk_{c}.tif") for c in chunk_ids]
    for s in srcs:
        assert [ci.name for ci in s.colorinterp[:3]] == ["red", "green", "blue"]
    arr, transform = merge(srcs)
    for s in srcs:
        s.close()
    img = Image.fromarray(np.transpose(arr[:3], (1, 2, 0)))
    return img, transform


def draw_view(chunk_ids, line, marks, out_name, labels=True, legend=True, credit=True,
              subtle=False, center=None, side=768, title=None):
    base, transform = mosaic(chunk_ids)
    inv = ~transform
    S = 2                                                   # supersample the overlay for smooth edges
    W, H = base.size
    ov = Image.new("RGBA", (W * S, H * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)

    def px(x, y):                                           # UTM -> overlay pixel
        lon, lat = to_ll.transform(x, y)
        c, r = inv * (lon, lat)
        return (c * S, r * S)

    def rings(geom):
        polys = geom.geoms if isinstance(geom, MultiPolygon) else [geom]
        for poly in polys:
            yield [px(*xy) for xy in poly.exterior.coords], [[px(*xy) for xy in i.coords] for i in poly.interiors]

    a = 0.55 if subtle else 1.0
    b100, b50 = line.buffer(100, quad_segs=32), line.buffer(50, quad_segs=32)
    for ext, holes in rings(b100):
        d.polygon(ext, fill=C_IN100 + (int(34 * a),))
    for ext, holes in rings(b50):
        d.polygon(ext, fill=C_IN50 + (int(40 * a),))
    for ext, _ in rings(b100):
        d.line(ext + [ext[0]], fill=C_IN100 + (int(230 * a),), width=3 * S)
    for ext, _ in rings(b50):
        d.line(ext + [ext[0]], fill=C_IN50 + (int(230 * a),), width=3 * S)
    route = [px(*xy) for xy in line.coords]
    d.line(route, fill=(13, 17, 23, int(200 * a)), width=11 * S, joint="curve")
    d.line(route, fill=C_ROUTE + (255,), width=6 * S, joint="curve")

    lab = font(15 * S, 700)
    placed = []
    for m in marks:
        cx, cy = px(m["p"].x, m["p"].y)
        if not (-40 <= cx <= W * S + 40 and -40 <= cy <= H * S + 40):
            continue
        col = C_IN50 if m["dist"] <= 50 else C_IN100 if m["dist"] <= 100 else C_BEYOND
        h = 12 * S                                          # a 12 m mark box at ~0.55 m/px
        d.rectangle([cx - h, cy - h, cx + h, cy + h], outline=(13, 17, 23, 200), width=6 * S)
        d.rectangle([cx - h, cy - h, cx + h, cy + h], outline=col + (255,), width=3 * S)
        if labels:
            tw = d.textlength(m["id"], font=lab)
            bw, bh = tw + 10 * S, 20 * S
            spots = [(cx - bw / 2, cy - h - 24 * S), (cx - bw / 2, cy + h + 4 * S),
                     (cx + h + 4 * S, cy - bh / 2), (cx - h - 4 * S - bw, cy - bh / 2)]
            for bx, by in spots:                            # first spot that clears earlier labels
                r = (bx, by, bx + bw, by + bh)
                if not any(r[0] < q[2] and q[0] < r[2] and r[1] < q[3] and q[1] < r[3] for q in placed):
                    break
            placed.append(r)
            d.rounded_rectangle(r, radius=4 * S, fill=(13, 17, 23, 225))
            d.text((r[0] + 5 * S, r[1] + 1 * S), m["id"], font=lab, fill=col + (255,))

    ov = ov.resize((W, H), Image.LANCZOS)
    img = Image.alpha_composite(base.convert("RGBA"), ov)
    dd = ImageDraw.Draw(img)

    if legend:
        items = [("Route (as supplied)", C_ROUTE, "line"), ("50 m buffer", C_IN50, "band"),
                 ("100 m buffer", C_IN100, "band"), ("Reviewer mark within 50 m", C_IN50, "box"),
                 ("Reviewer mark 50–100 m", C_IN100, "box"), ("Reviewer mark beyond 100 m", C_BEYOND, "box")]
        f = font(17, 600)
        pad, lh = 14, 28
        wmax = max(dd.textlength(t, font=f) for t, _, _ in items) + 60
        x0, y0 = 18, 18
        hh = pad * 2 + lh * len(items) + (30 if title else 0)
        dd.rounded_rectangle([x0, y0, x0 + wmax, y0 + hh], radius=10, fill=(13, 17, 23, 222))
        y = y0 + pad
        if title:
            dd.text((x0 + pad, y), title, font=font(18, 750), fill=(255, 255, 255))
            y += 30
        for t, c, kind in items:
            if kind == "line":
                dd.line([x0 + pad, y + 11, x0 + pad + 28, y + 11], fill=c, width=5)
            elif kind == "band":
                dd.rectangle([x0 + pad, y + 4, x0 + pad + 28, y + 18], fill=c + (90,), outline=c, width=2)
            else:
                dd.rectangle([x0 + pad + 6, y + 3, x0 + pad + 22, y + 19], outline=c, width=3)
            dd.text((x0 + pad + 40, y), t, font=f, fill=(235, 238, 242))
            y += lh

    if center is not None:                                  # square crop centred on a chainage
        c = line.interpolate(center)
        cx, cy = inv * to_ll.transform(c.x, c.y)
        x0 = int(min(max(cx - side / 2, 0), W - side))
        y0 = int(min(max(cy - side / 2, 0), H - side))
        img = img.crop((x0, y0, x0 + side, y0 + side))
    if credit:
        dd = ImageDraw.Draw(img)
        t = "Imagery © Google · marks and buffers: AfriScan"
        f = font(15, 500)
        tw = dd.textlength(t, font=f)
        w, h = img.size
        dd.rounded_rectangle([w - tw - 26, h - 34, w - 8, h - 8], radius=6, fill=(13, 17, 23, 200))
        dd.text((w - tw - 17, h - 31), t, font=f, fill=(235, 238, 242))

    OUT_IMG.mkdir(parents=True, exist_ok=True)
    img.convert("RGB").save(OUT_IMG / out_name, "JPEG", quality=90, optimize=True, progressive=True)
    return img.size


def main():
    line = load_route()
    marks = load_marks(line)
    segs = segments(line, marks)
    within = {b: sum(m["dist"] <= b for m in marks) for b in BUFFERS}

    # Dense stretch around km 5.0-6.3 (segments 5.0-5.5, 5.5-6.0 and 6.0-6.5 km)
    size_a = draw_view(["0080", "0081", "0103", "0104"], line, marks, "t9-km5-6.jpg",
                       title="T-9 route · km 5.0–6.3")
    # Wide view for the home hero: the route runs east-west through km 3.3-5.0
    size_h = draw_view(["0099", "0100", "0101", "0102", "0103", "0122", "0123", "0124", "0125", "0126"],
                       line, marks, "t9-hero.jpg", labels=False, legend=False, credit=False, subtle=True)
    # One square per rating, same scale: high (5.5-6.0 km), medium (8.5-9.0 km), low (1.0-1.5 km)
    size_r = []
    for name, chunks, chain in (("t9-rating-high.jpg", ["0080", "0081", "0103", "0104"], 5750),
                                ("t9-rating-medium.jpg", ["0017", "0018", "0019", "0040", "0041", "0042"], 8990),
                                ("t9-rating-low.jpg", ["0015", "0016", "0038", "0039"], 8000)):
        size_r.append(draw_view(chunks, line, marks, name, labels=False, legend=False, center=chain))

    in_view = [m for m in marks if 35.10132 <= m["lon"] <= 35.10956 and -21.74037 <= m["lat"] <= -21.73271]
    data = {
        "route": "T-9 replacement pipeline",
        "province": "Inhambane, Mozambique",
        "route_length_km": round(line.length / 1000, 2),
        "buffers_m": list(BUFFERS),
        "method": "Reviewer-marked structures (manual review), no automatic detection in this sample",
        "imagery": "Google satellite basemap, shown for illustration; capture date not supplied by the provider",
        "marks_total": len(marks),
        "within_m": {str(k): v for k, v in within.items()},
        "segment_m": SEGMENT,
        "rating_rule": {"high_over": HIGH_OVER, "buffer_m": max(BUFFERS)},
        "segments": segs,
        "segments_by_rating": {r: sum(s["rating"] == r for s in segs) for r in ("high", "medium", "low")},
        "register_excerpt": [{"id": m["id"], "km": round(m["chain"] / 1000, 2), "distance_m": round(m["dist"]),
                              "band": m["band"]} for m in in_view],
        "images": {"t9-km5-6": size_a, "t9-hero": size_h,
                   "t9-rating-high": size_r[0], "t9-rating-medium": size_r[1], "t9-rating-low": size_r[2]},
        "source_job": "corridor_909497cd (manual review, complete)",
    }
    OUT_DATA.parent.mkdir(parents=True, exist_ok=True)
    OUT_DATA.write_text(json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({k: data[k] for k in ("marks_total", "within_m", "segments_by_rating")}, ensure_ascii=False))
    print("register excerpt:", len(in_view), "marks in the km 5-6 view")


if __name__ == "__main__":
    main()
