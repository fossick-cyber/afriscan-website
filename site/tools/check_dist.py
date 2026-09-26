#!/usr/bin/env python3
"""Independent checks on a built dist/: hreflang, canonicals, sitemaps, links, JSON-LD, orphans.

Usage: check_dist.py [dist_dir]   (exit 1 on any error)

It re-reads the generated HTML with the standard library only, so it does not share code (or bugs)
with build.py.
"""
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from html.parser import HTMLParser
from urllib.parse import urlsplit, unquote

ORIGIN = "https://afri-scan.com"
CODES = {"en", "x-default", "en-MZ", "pt-MZ", "pt", "en-ZA", "en-NG"}
LANG_BY_PREFIX = [("/mz/pt/", "pt-MZ"), ("/mz/", "en-MZ"), ("/za/", "en-ZA"), ("/ng/", "en-NG"), ("/", "en-GB")]
OWN_CODE = {"pt-MZ": "pt-MZ", "en-MZ": "en-MZ", "en-ZA": "en-ZA", "en-NG": "en-NG", "en-GB": "en"}


class Page(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.lang = None
        self.canonical = []
        self.alternates = []
        self.robots = ""
        self.og_url = None
        self.ids = set()
        self.links = []          # (tag, attr, value, region)
        self.jsonld = []
        self._in_ld = False
        self._ld = []
        self.region = []         # stack of header/footer/nav/main
        self.h1 = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "html":
            self.lang = a.get("lang")
        if "id" in a:
            self.ids.add(a["id"])
        if tag == "a" and "name" in a:
            self.ids.add(a["name"])
        if tag in ("header", "footer", "nav", "main"):
            self.region.append(tag)
        if tag == "h1":
            self.h1 += 1
        reg = self.region[0] if self.region else "body"
        if tag == "link":
            rel = (a.get("rel") or "").lower()
            if rel == "canonical":
                self.canonical.append(a.get("href"))
            elif rel == "alternate" and a.get("hreflang"):
                self.alternates.append((a["hreflang"], a.get("href")))
            elif a.get("href"):
                self.links.append((tag, "href", a["href"], "head"))
        elif tag == "meta":
            if (a.get("name") or "").lower() == "robots":
                self.robots = a.get("content") or ""
            if a.get("property") == "og:url":
                self.og_url = a.get("content")
            if a.get("property") in ("og:image",) or a.get("name") == "twitter:image":
                self.links.append((tag, "content", a.get("content"), "head"))
        elif tag == "script":
            if a.get("type") == "application/ld+json":
                self._in_ld = True
                self._ld = []
            elif a.get("src"):
                self.links.append((tag, "src", a["src"], reg))
        else:
            for attr in ("href", "src", "poster", "action"):
                if a.get(attr):
                    self.links.append((tag, attr, a[attr], reg))
            for attr in ("srcset", "imagesrcset"):
                if a.get(attr):
                    for part in a[attr].split(","):
                        u = part.strip().split(" ")[0]
                        if u:
                            self.links.append((tag, attr, u, reg))

    def handle_endtag(self, tag):
        if tag == "script" and self._in_ld:
            self._in_ld = False
            self.jsonld.append("".join(self._ld))
        if tag in ("header", "footer", "nav", "main") and self.region:
            # pop the innermost matching region
            for i in range(len(self.region) - 1, -1, -1):
                if self.region[i] == tag:
                    del self.region[i]
                    break

    def handle_data(self, data):
        if self._in_ld:
            self._ld.append(data)


def url_for(rel):
    rel = rel.replace(os.sep, "/")
    if rel == "index.html":
        return "/"
    if rel.endswith("/index.html"):
        return "/" + rel[: -len("index.html")]
    return "/" + rel[: -len(".html")]


def file_for(dist, path):
    """Map a clean URL path to a file in dist the way Cloudflare Pages does."""
    path = unquote(path)
    if path.endswith("/"):
        cand = [path + "index.html"]
    else:
        cand = [path, path + ".html", path + "/index.html"]
    for c in cand:
        f = os.path.join(dist, c.lstrip("/"))
        if os.path.isfile(f):
            return f
    return None


def main():
    dist = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "dist")
    errors, warns = [], []
    pages = {}
    for root, _, files in os.walk(dist):
        for fn in files:
            if fn.endswith(".html"):
                rel = os.path.relpath(os.path.join(root, fn), dist)
                p = Page()
                with open(os.path.join(root, fn), encoding="utf-8") as fh:
                    p.feed(fh.read())
                pages[url_for(rel)] = p

    def is_404(u):
        return u.endswith("/404")

    indexable = {u for u, p in pages.items() if "noindex" not in p.robots.lower() and not is_404(u)}

    # --- canonical, og:url, lang, H1 ---
    for u, p in sorted(pages.items()):
        if is_404(u):
            continue
        noindex = u not in indexable
        want = ORIGIN + u
        if not noindex:
            if p.canonical != [want]:
                errors.append(f"{u}: canonical {p.canonical} != [{want}]")
            if p.og_url != want:
                errors.append(f"{u}: og:url {p.og_url} != {want}")
        lang = next(l for pre, l in LANG_BY_PREFIX if u.startswith(pre))
        if p.lang != lang:
            errors.append(f"{u}: <html lang={p.lang}> expected {lang}")
        if p.h1 != 1:
            errors.append(f"{u}: {p.h1} <h1> elements")

    # --- hreflang ---
    clusters = 0
    for u, p in sorted(pages.items()):
        if not p.alternates:
            continue
        if u not in indexable:
            errors.append(f"{u}: noindex page carries hreflang")
        codes = [c for c, _ in p.alternates]
        if len(codes) != len(set(codes)):
            errors.append(f"{u}: duplicate hreflang codes {codes}")
        bad = set(codes) - CODES
        if bad:
            errors.append(f"{u}: unknown hreflang codes {bad}")
        amap = dict(p.alternates)
        if "x-default" not in amap:
            errors.append(f"{u}: hreflang cluster without x-default")
        if ORIGIN + u not in amap.values():
            errors.append(f"{u}: hreflang does not reference itself")
        own = OWN_CODE[next(l for pre, l in LANG_BY_PREFIX if u.startswith(pre))]
        if amap.get(own) != ORIGIN + u:
            errors.append(f"{u}: own code {own} points to {amap.get(own)}")
        if "pt-MZ" in amap and amap.get("pt") != amap["pt-MZ"]:
            errors.append(f"{u}: pt catch-all {amap.get('pt')} != pt-MZ {amap['pt-MZ']}")
        if "en" in amap and "x-default" in amap and "en" in amap and amap["en"] != amap["x-default"]:
            errors.append(f"{u}: en and x-default differ")
        for code, href in p.alternates:
            if not href.startswith(ORIGIN + "/"):
                errors.append(f"{u}: hreflang {code} not absolute: {href}")
                continue
            tu = href[len(ORIGIN):]
            if tu not in pages:
                errors.append(f"{u}: hreflang {code} -> {tu} is not a built page")
                continue
            if tu not in indexable:
                errors.append(f"{u}: hreflang {code} -> {tu} is noindex")
            tp = pages[tu]
            if tp.canonical != [href]:
                errors.append(f"{u}: hreflang {code} -> {tu} whose canonical is {tp.canonical}")
            if sorted(tp.alternates) != sorted(p.alternates):
                errors.append(f"{u}: hreflang set differs on {tu} (not reciprocal)")
        # section check: every alternate URL's section matches its code
        for code, href in p.alternates:
            tu = href[len(ORIGIN):]
            sec = OWN_CODE[next(l for pre, l in LANG_BY_PREFIX if tu.startswith(pre))]
            if code not in ("x-default", "pt") and code != sec:
                errors.append(f"{u}: hreflang {code} points into section {sec}: {tu}")
        clusters += 1

    # --- sitemaps ---
    ns = {"s": "http://www.sitemaps.org/schemas/sitemap/0.9"}
    in_sitemap = defaultdict(list)
    try:
        idx = ET.parse(os.path.join(dist, "sitemap.xml")).getroot()
        subs = [e.text for e in idx.findall("s:sitemap/s:loc", ns)]
        if not subs:
            errors.append("sitemap.xml has no child sitemaps")
        for sm in subs:
            f = os.path.join(dist, sm[len(ORIGIN) + 1:])
            if not os.path.isfile(f):
                errors.append(f"sitemap index lists missing {sm}")
                continue
            root = ET.parse(f).getroot()
            for url in root.findall("s:url", ns):
                loc = url.find("s:loc", ns).text
                lm = url.find("s:lastmod", ns)
                if lm is None or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", lm.text or ""):
                    warns.append(f"{os.path.basename(f)}: {loc} lastmod {lm.text if lm is not None else None}")
                in_sitemap[loc].append(os.path.basename(f))
                if url.find("s:priority", ns) is not None or url.find("s:changefreq", ns) is not None:
                    warns.append(f"{loc}: priority/changefreq present")
    except Exception as e:  # noqa: BLE001
        errors.append(f"sitemap parse failed: {e}")
    for loc, where in in_sitemap.items():
        if len(where) > 1:
            errors.append(f"{loc} listed {len(where)} times in sitemaps {where}")
        tu = loc[len(ORIGIN):] if loc.startswith(ORIGIN) else None
        if tu not in pages:
            errors.append(f"sitemap URL {loc} is not a built page")
        elif tu not in indexable:
            errors.append(f"sitemap URL {loc} is noindex")
        else:
            sec = "mz" if tu.startswith("/mz/") else "za" if tu.startswith("/za/") else "ng" if tu.startswith("/ng/") else "global"
            if where[0] != f"sitemap-{sec}.xml":
                errors.append(f"{loc} is in {where[0]}, expected sitemap-{sec}.xml")
    for u in sorted(indexable):
        if ORIGIN + u not in in_sitemap:
            errors.append(f"indexable page {u} missing from sitemaps")

    # --- robots.txt ---
    rb = open(os.path.join(dist, "robots.txt"), encoding="utf-8").read()
    if f"Sitemap: {ORIGIN}/sitemap.xml" not in rb:
        errors.append("robots.txt lacks the sitemap line")

    # --- links and anchors ---
    inbound = defaultdict(set)          # target url -> set of source urls (body links only)
    inbound_nav = defaultdict(set)      # from header/footer/nav
    checked = 0
    for u, p in sorted(pages.items()):
        for tag, attr, val, reg in p.links:
            if val is None:
                continue
            if val.startswith(("mailto:", "tel:", "javascript:", "data:", "#")):
                if val.startswith("#") and len(val) > 1 and unquote(val[1:]) not in p.ids:
                    errors.append(f"{u}: in-page anchor {val} missing")
                continue
            parts = urlsplit(val)
            if parts.scheme in ("http", "https") and parts.netloc not in ("afri-scan.com", "www.afri-scan.com"):
                continue
            if parts.scheme in ("http", "https") and parts.netloc == "www.afri-scan.com":
                errors.append(f"{u}: link via www host {val}")
            if parts.scheme == "" and not val.startswith("/"):
                errors.append(f"{u}: relative link {val}")
                continue
            if parts.scheme == "http":
                errors.append(f"{u}: http link {val}")
            path = parts.path or "/"
            checked += 1
            if path.endswith(".html") and tag == "a":
                errors.append(f"{u}: .html link {val}")
            f = file_for(dist, path)
            if f is None:
                # redirects file counts as resolving only for documented aliases; flag anyway
                errors.append(f"{u}: broken {tag}[{attr}] {val}")
                continue
            if tag == "a" and f.endswith(".html"):
                tu = url_for(os.path.relpath(f, dist))
                if tu != path:
                    errors.append(f"{u}: link {val} is not the clean URL {tu}")
                if parts.fragment:
                    frag = unquote(parts.fragment)
                    if frag not in pages[tu].ids:
                        errors.append(f"{u}: anchor #{frag} missing on {tu}")
                if tu != u:
                    (inbound_nav if reg in ("header", "footer", "nav") else inbound)[tu].add(u)

    orphans = []
    for u in sorted(indexable):
        if u == "/":
            continue
        if not inbound[u] and not inbound_nav[u]:
            orphans.append(u)
            errors.append(f"{u}: no inbound links from any page")
    body_only = [u for u in sorted(indexable) if u != "/" and not inbound[u]]

    # --- JSON-LD ---
    ld_blocks = 0
    types = defaultdict(int)
    for u, p in sorted(pages.items()):
        for block in p.jsonld:
            ld_blocks += 1
            try:
                data = json.loads(block)
            except Exception as e:  # noqa: BLE001
                errors.append(f"{u}: JSON-LD does not parse: {e}")
                continue
            txt = json.dumps(data)
            if re.search(r'"@type":\s*"(Offer|AggregateOffer|PriceSpecification)"', txt) or '"price' in txt.lower():
                errors.append(f"{u}: JSON-LD carries an offer or price")
            if data.get("@context") not in ("https://schema.org", "https://schema.org/"):
                errors.append(f"{u}: JSON-LD @context {data.get('@context')}")
            nodes = data.get("@graph", [data])
            for n in nodes:
                t = n.get("@type")
                for tt in (t if isinstance(t, list) else [t]):
                    types[tt] += 1
                if "url" in n and isinstance(n["url"], str) and n["url"].startswith(ORIGIN):
                    tu = n["url"][len(ORIGIN):].split("#")[0]
                    if tu not in pages:
                        errors.append(f"{u}: JSON-LD url {n['url']} not built")
                if t == "BreadcrumbList":
                    for it in n.get("itemListElement", []):
                        iu = it.get("item", "")
                        if iu.startswith(ORIGIN) and iu[len(ORIGIN):] not in pages:
                            errors.append(f"{u}: breadcrumb item {iu} not built")
                if t == "FAQPage" and not n.get("mainEntity"):
                    errors.append(f"{u}: FAQPage without questions")

    print(f"pages {len(pages)}, indexable {len(indexable)}, hreflang pages {clusters}, "
          f"sitemap URLs {len(in_sitemap)}, links checked {checked}, JSON-LD blocks {ld_blocks}")
    print("JSON-LD types:", ", ".join(f"{k} {v}" for k, v in sorted(types.items(), key=lambda x: -x[1])))
    if body_only:
        print(f"reachable only from header/footer menus ({len(body_only)}):", " ".join(body_only))
    for w in warns:
        print("WARN ", w)
    for e in errors:
        print("ERROR", e)
    print(f"{len(errors)} errors, {len(warns)} warnings")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
