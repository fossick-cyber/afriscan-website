"""Social cards (1200x630) and the favicon / logo set, drawn with Pillow in the site's own font.

Cards are text on the dark brand ground with a drawn corridor motif (route line and buffer bands):
no client imagery, no basemap imagery, no flags or regulator logos.
"""
import hashlib
import json
import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

INK, INK2, ACCENT, TEAL, TEXT, DIM = "#0d1117", "#0b3954", "#e8672f", "#5eead4", "#f3f5f7", "#aab4be"
CARD_VERSION = "og-v3"

# The mark: an orange "A" on the dark rounded square. Coordinates on a 64-unit grid, shared by the
# SVG favicon and the PNG/ICO renders so they match.
A_OUTER = [(32, 11), (50, 53), (41.5, 53), (37.9, 44), (26.1, 44), (22.5, 53), (14, 53)]
A_INNER = [(32, 24.5), (35.4, 36.5), (28.6, 36.5)]
SVG_MARK = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">'
            '<rect width="64" height="64" rx="12" fill="#0d1117"/>'
            '<path fill="#e8672f" fill-rule="evenodd" d="M32 11 50 53h-8.5l-3.6-9H26.1l-3.6 9H14Z'
            'M32 24.5l-3.4 12h6.8Z"/></svg>')


def _font(path, size, weight):
    f = ImageFont.truetype(str(path), size)
    try:
        f.set_variation_by_axes([weight])
    except Exception:
        pass
    return f


def mark_png(size):
    S = 4
    im = Image.new("RGBA", (size * S, size * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    k = size * S / 64
    d.rounded_rectangle([0, 0, size * S - 1, size * S - 1], radius=12 * k, fill=INK)
    d.polygon([(x * k, y * k) for x, y in A_OUTER], fill=ACCENT)
    d.polygon([(x * k, y * k) for x, y in A_INNER], fill=INK)
    return im.resize((size, size), Image.LANCZOS)


def write_icons(dist: Path, theme="#0d1117"):
    (dist / "favicon.svg").write_text(SVG_MARK, encoding="utf-8")
    mark_png(180).convert("RGB").save(dist / "apple-touch-icon.png", optimize=True)
    mark_png(192).save(dist / "icon-192.png", optimize=True)
    mark_png(512).save(dist / "icon-512.png", optimize=True)
    mark_png(48).save(dist / "favicon.ico", sizes=[(16, 16), (32, 32), (48, 48)])
    brand = dist / "assets/brand"
    brand.mkdir(parents=True, exist_ok=True)
    mark_png(512).save(brand / "afriscan-logo-512.png", optimize=True)
    (brand / "afriscan-mark.svg").write_text(SVG_MARK, encoding="utf-8")
    (dist / "site.webmanifest").write_text(json.dumps({
        "name": "AfriScan by Afridrone", "short_name": "AfriScan", "start_url": "/", "display": "browser",
        "theme_color": theme, "background_color": theme,
        "icons": [{"src": "/icon-192.png", "sizes": "192x192", "type": "image/png"},
                  {"src": "/icon-512.png", "sizes": "512x512", "type": "image/png"}]}, indent=1), encoding="utf-8")


def _corridor(img):
    """Route line with 50/100 m style buffer bands, bottom right, low contrast."""
    W, H = img.size
    S = 2
    ov = Image.new("RGBA", (W * S, H * S), (0, 0, 0, 0))
    d = ImageDraw.Draw(ov)
    pts = [(1250, 90), (1060, 250), (980, 330), (900, 470), (760, 700)]
    pts = [(x * S, y * S) for x, y in pts]
    d.line(pts, fill=(245, 158, 11, 30), width=190 * S, joint="curve")
    d.line(pts, fill=(239, 68, 68, 40), width=96 * S, joint="curve")
    d.line(pts, fill=(232, 103, 47, 235), width=7 * S, joint="curve")
    for x, y, c in ((1080, 180, TEAL), (1010, 360, "#f59e0b"), (905, 400, "#ef4444"), (1000, 470, "#f59e0b"),
                    (845, 560, "#ef4444"), (1130, 330, TEAL)):
        r = 13 * S
        d.rectangle([x * S - r, y * S - r, x * S + r, y * S + r], outline=c, width=4 * S)
    ov = ov.resize((W, H), Image.LANCZOS)
    img.alpha_composite(ov)


def card(out_dir: Path, url_path: str, headline: str, subline: str, region: str, font_path: Path,
         cache_dir: Path) -> str:
    key = hashlib.sha256("|".join([CARD_VERSION, headline, subline, region]).encode()).hexdigest()[:10]
    stem = url_path.strip("/").replace("/", "-") or "home"
    name = f"{stem}-{key}.jpg"
    cached = cache_dir / name
    cache_dir.mkdir(parents=True, exist_ok=True)
    if not cached.exists():
        W, H = 1200, 630
        img = Image.new("RGBA", (W, H), INK)
        grad = Image.linear_gradient("L").rotate(90).resize((W, H))
        blue = Image.new("RGBA", (W, H), INK2)
        img = Image.composite(blue, img, grad.point(lambda v: int(v * 0.55)))
        _corridor(img)
        d = ImageDraw.Draw(img)
        d.rectangle([0, 0, 14, H], fill=ACCENT)
        wm = _font(font_path, 46, 800)
        d.text((72, 58), "Afri", font=wm, fill=TEXT)
        d.text((72 + d.textlength("Afri", font=wm), 58), "Scan", font=wm, fill=ACCENT)
        by = _font(font_path, 24, 500)
        d.text((72 + d.textlength("AfriScan", font=wm) + 14, 76), "by Afridrone", font=by, fill=DIM)
        if region:
            rf = _font(font_path, 24, 700)
            tw = d.textlength(region.upper(), font=rf)
            d.rounded_rectangle([72, 138, 72 + tw + 28, 180], radius=8, fill=(22, 27, 34), outline=(60, 70, 82))
            d.text((86, 146), region.upper(), font=rf, fill=TEAL)
        hf = _font(font_path, 58, 800)
        lines = textwrap.wrap(headline, 26)[:4]
        y = 236 if len(lines) <= 3 else 212
        for line in lines:
            d.text((72, y), line, font=hf, fill=TEXT)
            y += 70
        sf = _font(font_path, 28, 500)
        for line in textwrap.wrap(subline, 52)[:2]:
            d.text((72, y + 18), line, font=sf, fill=DIM)
            y += 40
        img.convert("RGB").save(cached, "JPEG", quality=84, optimize=True, progressive=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / name
    if not target.exists():
        target.write_bytes(cached.read_bytes())
    return name
