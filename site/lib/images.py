"""Responsive images: masters in site/images/ -> AVIF, WebP and JPEG at several widths.

Output names carry a hash of the master and the encoder settings, so /assets/* can be cached as
immutable. Encoded files are kept in site/.cache/img/ between builds.
"""
import hashlib
import shutil
from html import escape
from pathlib import Path

from PIL import Image

WIDTHS = (480, 800, 1200, 1600, 2000, 2400)
QUALITY = {"avif": 52, "webp": 74, "jpg": 80}
VERSION = "img-v2"


class ImageError(Exception):
    pass


class Images:
    def __init__(self, src_dir: Path, cache_dir: Path, dist_dir: Path, url_prefix="/assets/img"):
        self.src_dir, self.cache_dir, self.dist_dir = src_dir, cache_dir, dist_dir
        self.url_prefix = url_prefix
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.done = {}
        self.used = []            # (page_url, jpg_url) for the image sitemap

    def master(self, name):
        for ext in (".jpg", ".jpeg", ".png", ".webp"):
            p = self.src_dir / f"{name}{ext}"
            if p.exists():
                return p
        raise ImageError(f"image '{name}' not found in site/images/ (looked for .jpg, .png, .webp)")

    def variants(self, name, max_width=None):
        """Encode (or reuse) every variant; returns dict(fmt -> [(w, url)]), size."""
        key = (name, max_width)
        if key in self.done:
            return self.done[key]
        src = self.master(name)
        digest = hashlib.sha256(src.read_bytes() + VERSION.encode()).hexdigest()[:8]
        with Image.open(src) as im:
            im = im.convert("RGB")
            W, H = im.size
            top = min(W, max_width or W, WIDTHS[-1])
            widths = [w for w in WIDTHS if w < top] + [top]
            out = {"avif": [], "webp": [], "jpg": []}
            stem = name.replace("/", "-")
            for w in widths:
                h = round(H * w / W)
                resized = None
                for fmt in ("avif", "webp", "jpg"):
                    fname = f"{stem}-{digest}-{w}.{fmt}"
                    cached = self.cache_dir / fname
                    if not cached.exists():
                        if resized is None:
                            resized = im if w == W else im.resize((w, h), Image.LANCZOS)
                        tmp = cached.with_suffix(".tmp")
                        if fmt == "avif":
                            resized.save(tmp, "AVIF", quality=QUALITY["avif"], speed=6)
                        elif fmt == "webp":
                            resized.save(tmp, "WEBP", quality=QUALITY["webp"], method=6)
                        else:
                            resized.save(tmp, "JPEG", quality=QUALITY["jpg"], optimize=True, progressive=True)
                        tmp.rename(cached)
                    target = self.dist_dir / fname
                    if not target.exists():
                        target.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copyfile(cached, target)
                    out[fmt].append((w, f"{self.url_prefix}/{fname}"))
        result = (out, (W, H), top)
        self.done[key] = result
        return result

    def picture(self, name, alt, sizes="100vw", priority=False, cls="", max_width=None, page_url=None):
        if alt is None:
            raise ImageError(f"image '{name}': alt text is required (use alt=\"\" only for decoration)")
        out, (W, H), top = self.variants(name, max_width)
        h = round(H * top / W)
        srcset = {f: ", ".join(f"{u} {w}w" for w, u in v) for f, v in out.items()}
        mid = next((u for w, u in out["jpg"] if w >= 1200), out["jpg"][-1][1])
        if page_url:
            self.used.append((page_url, out["jpg"][-1][1]))
        load = 'fetchpriority="high" decoding="async"' if priority else 'loading="lazy" decoding="async"'
        klass = f' class="{escape(cls)}"' if cls else ""
        return (f'<picture><source type="image/avif" srcset="{srcset["avif"]}" sizes="{escape(sizes)}">'
                f'<source type="image/webp" srcset="{srcset["webp"]}" sizes="{escape(sizes)}">'
                f'<img src="{mid}" srcset="{srcset["jpg"]}" sizes="{escape(sizes)}" width="{top}" height="{h}" '
                f'alt="{escape(alt)}"{klass} {load}></picture>')

    def largest(self, name, fmt="jpg"):
        out, _, _ = self.variants(name)
        return out[fmt][-1][1]
