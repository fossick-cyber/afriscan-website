#!/usr/bin/env python3
"""Independent checks on a built dist/: hreflang, canonicals, sitemaps, links, JSON-LD, orphans, and
that nothing published points into a draft country section.

    /opt/favhousecheck/.venv/bin/python3 site/tools/check_dist.py [DIST] [--drafts] [--data DIR]

It re-reads the generated files with the standard library (plus PyYAML for the section list in
data/locales.yaml and data/countries.yaml), so it does not share code, or bugs, with build.py.
Exit 1 on any error.
"""
import argparse
import json
import os
import re
import sys
import xml.etree.ElementTree as ET
from collections import defaultdict
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlsplit, unquote

try:
    import yaml
except ImportError:  # pragma: no cover
    sys.exit("check_dist.py needs PyYAML: run it with /opt/favhousecheck/.venv/bin/python3")

ORIGIN = "https://afri-scan.com"
DRAFT_BANNER = "Draft: not published"
# elements whose links are reported by what holds them (class on the element or an ancestor)
HOLDERS = [("region", "country menu"), ("country-sites", "country-sites button"),
           ("countries-directory", "/countries entry")]


class Sections:
    """data/locales.yaml, read independently of build.py."""

    def __init__(self, data_dir):
        self.loc = yaml.safe_load(Path(data_dir, "locales.yaml").read_text(encoding="utf-8"))
        cpath = Path(data_dir, "countries.yaml")
        self.countries = yaml.safe_load(cpath.read_text(encoding="utf-8")) if cpath.exists() else {}
        self.live = [k for k, l in self.loc.items() if l.get("status", "live") == "live"]
        self.draft = [k for k, l in self.loc.items() if l.get("status", "live") != "live"]
        # longest prefix first, so /mz/pt/ wins over /mz/
        self.by_prefix = sorted(((l["prefix"] + "/", k) for k, l in self.loc.items()), key=lambda x: -len(x[0]))
        self.own = {k: l["hreflang"][0] for k, l in self.loc.items()}
        self.catch = {c: k for k, l in self.loc.items() for c in l["hreflang"][1:] if c != "x-default"}
        self.live_codes = {c for k in self.live for c in self.loc[k]["hreflang"]} | {"x-default"}
        self.draft_langs = {self.loc[k]["lang"] for k in self.draft}
        live_cc = {self.loc[k]["country"] for k in self.live if self.loc[k].get("country")}
        self.draft_cc = {self.loc[k]["country"] for k in self.draft if self.loc[k].get("country")} - live_cc
        self.draft_names = set()
        for k in self.draft:
            cc = self.loc[k].get("country")
            if cc in self.draft_cc:
                self.draft_names |= {self.loc[k]["country_name"], (self.countries.get(cc.lower()) or {}).get("name")}
        self.draft_names.discard(None)
        live_maps = {self.loc[k]["sitemap"] for k in self.live}
        self.draft_maps = {self.loc[k]["sitemap"] for k in self.draft} - live_maps

    def name_owner(self, name):
        return next((k for k in self.draft if name in (self.loc[k]["country_name"], (self.countries.get(
            (self.loc[k].get("country") or "").lower()) or {}).get("name"))), "?")

    def of(self, path):
        """Section key of a URL path (or None for a path outside the site)."""
        if not path.startswith("/"):
            return None
        for pre, k in self.by_prefix:
            if path.startswith(pre) or path + "/" == pre:
                return k
        return None

    def is_draft(self, path):
        k = self.of(path)
        return k if k in self.draft else None

    def lang(self, path):
        return self.loc[self.of(path)]["lang"]


class Page(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.lang = None
        self.canonical = []
        self.alternates = []
        self.robots = ""
        self.og_url = None
        self.ids = set()
        self.dup_ids = []
        self.links = []          # (tag, attr, value, region, holder)
        self.jsonld = []
        self._in_ld = False
        self._ld = []
        self.region = []         # stack of header/footer/nav/main
        self.stack = []          # (tag, holder) for open elements
        self.h1 = 0
        self.text = []

    VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "html":
            self.lang = a.get("lang")
        if "id" in a:
            if a["id"] in self.ids:
                self.dup_ids.append(a["id"])
            self.ids.add(a["id"])
        if tag == "a" and "name" in a:
            self.ids.add(a["name"])
        if tag in ("header", "footer", "nav", "main"):
            self.region.append(tag)
        if tag == "h1":
            self.h1 += 1
        classes = (a.get("class") or "").split()
        own = next((what for cls, what in HOLDERS if cls in classes), None)
        holder = own or next((h for _, h in reversed(self.stack) if h), None)
        if tag not in self.VOID:
            self.stack.append((tag, own))
        reg = self.region[0] if self.region else "body"
        holder = holder or (f"{reg} link" if reg in ("header", "footer") else "internal link")
        if tag == "link":
            rel = (a.get("rel") or "").lower()
            if rel == "canonical":
                self.canonical.append(a.get("href"))
            elif rel == "alternate" and a.get("hreflang"):
                self.alternates.append((a["hreflang"], a.get("href")))
            elif a.get("href"):
                self.links.append((tag, "href", a["href"], "head", "head link"))
        elif tag == "meta":
            if (a.get("name") or "").lower() == "robots":
                self.robots = a.get("content") or ""
            if a.get("property") == "og:url":
                self.og_url = a.get("content")
            if a.get("property") in ("og:image",) or a.get("name") == "twitter:image":
                self.links.append((tag, "content", a.get("content"), "head", "meta tag"))
        elif tag == "script":
            if a.get("type") == "application/ld+json":
                self._in_ld = True
                self._ld = []
            elif a.get("src"):
                self.links.append((tag, "src", a["src"], reg, holder))
        else:
            for attr in ("href", "src", "poster", "action"):
                if a.get(attr):
                    self.links.append((tag, attr, a[attr], reg, holder))
            for attr in ("srcset", "imagesrcset"):
                if a.get(attr):
                    for part in a[attr].split(","):
                        u = part.strip().split(" ")[0]
                        if u:
                            self.links.append((tag, attr, u, reg, holder))

    def handle_endtag(self, tag):
        if tag == "script" and self._in_ld:
            self._in_ld = False
            self.jsonld.append("".join(self._ld))
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                del self.stack[i:]
                break
        if tag in ("header", "footer", "nav", "main") and self.region:
            # pop the innermost matching region
            for i in range(len(self.region) - 1, -1, -1):
                if self.region[i] == tag:
                    del self.region[i]
                    break

    def handle_data(self, data):
        if self._in_ld:
            self._ld.append(data)
        else:
            self.text.append(data)


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


def site_path(v):
    """The site path of a link value (root-relative or on ORIGIN), else None."""
    v = (v or "").strip()
    if v.startswith(ORIGIN):
        v = v[len(ORIGIN):] or "/"
    if not v.startswith("/") or v.startswith("//"):
        return None
    return re.split(r"[?#]", v)[0]


def walk_json(v, key=None):
    if isinstance(v, dict):
        for k, x in v.items():
            yield from walk_json(x, k)
        if v.get("@type") == "Country":
            yield ("country", v.get("name"))
    elif isinstance(v, list):
        for x in v:
            yield from walk_json(x, key)
    elif isinstance(v, str):
        yield (key, v)


def check_drafts(dist, S, pages, in_sitemap, sitemap_files, drafts, errors):
    """Nothing published may point into a draft section (a --drafts preview: draft pages are marked)."""
    def leak(where, what, value, lk):
        errors.append(f"[draft-leak] {where}: {what} -> {value} (draft section {lk})")

    # every file under a draft section's prefix
    if not drafts:
        for root, _, files in os.walk(dist):
            for fn in files:
                rel = "/" + os.path.relpath(os.path.join(root, fn), dist).replace(os.sep, "/")
                if lk := S.is_draft(rel):
                    leak("dist", "file in a draft section", rel, lk)
    for u, p in sorted(pages.items()):
        own = S.is_draft(u)
        banner = DRAFT_BANNER in " ".join(p.text)
        if own or banner:
            if drafts:
                if "noindex" not in p.robots.lower():
                    errors.append(f"[draft-page] {u}: a draft page must be noindex")
                if not banner:
                    errors.append(f"[draft-page] {u}: a draft page must show the “{DRAFT_BANNER}” banner")
                if p.alternates:
                    errors.append(f"[draft-page] {u}: a draft page carries hreflang")
            else:
                errors.append(f"[draft-page] {u}: a normal build must not contain draft pages")
            if own:
                continue
        # links from a published page (a --drafts preview links its drafts from every menu by design)
        if not drafts:
            for tag, attr, val, reg, holder in p.links:
                if (path := site_path(val)) and (lk := S.is_draft(path)):
                    leak(u, holder, val, lk)
        for code, href in p.alternates:
            lk = S.is_draft(site_path(href) or "") or next((k for k in S.draft if code in S.loc[k]["hreflang"]), None)
            if lk:
                leak(u, "hreflang alternate", f"{code} {href}", lk)
        for block in p.jsonld:
            try:
                data = json.loads(block)
            except ValueError:
                continue
            for key, v in walk_json(data):
                if key == "country" and v in S.draft_names:
                    leak(u, "JSON-LD country", v, S.name_owner(v))
                elif isinstance(v, str) and (path := site_path(v)) and (lk := S.is_draft(path)):
                    leak(u, "JSON-LD reference", v, lk)
                elif key == "inLanguage" and v in S.draft_langs:
                    leak(u, "JSON-LD inLanguage", v, next(k for k in S.draft if S.loc[k]["lang"] == v))
    for loc, where in in_sitemap.items():
        if (path := site_path(loc)) and (lk := S.is_draft(path)):
            leak(where[0], "sitemap entry", loc, lk)
    for name in sitemap_files:
        if name in S.draft_maps:
            leak("sitemap.xml", "sitemap of a draft section", name,
                 next(k for k in S.draft if S.loc[k]["sitemap"] == name))
    red = os.path.join(dist, "_redirects")
    if os.path.isfile(red) and not drafts:
        for line in open(red, encoding="utf-8"):
            parts = line.split()
            if len(parts) >= 2 and not line.startswith("#"):
                for v in parts[:2]:
                    if (path := site_path(v.rstrip("*"))) and (lk := S.is_draft(path)):
                        leak("_redirects", "redirect rule", line.strip(), lk)
    hdr = os.path.join(dist, "_headers")
    if os.path.isfile(hdr) and not drafts:
        for line in open(hdr, encoding="utf-8"):
            if line.startswith("/") and (lk := S.is_draft(line.strip().rstrip("*"))):
                leak("_headers", "_headers rule", line.strip(), lk)
    # the geo banner never suggests a draft section, in any build
    for f in sorted(Path(dist, "assets/js").glob("region.*.js")):
        js = f.read_text(encoding="utf-8")
        m = re.search(r"const DATA = (\{.*?\});\n", js, re.S)
        if not m:
            errors.append(f"/assets/js/{f.name}: no DATA object found (the banner check cannot run)")
            continue
        data = json.loads(m.group(1))
        for cc, site in (data.get("sites") or {}).items():
            if cc in S.draft_cc:
                leak(f"/assets/js/{f.name}", "geo banner country", cc,
                     next(k for k in S.draft if S.loc[k].get("country") == cc))
            for code, home, _ in site.get("links", []):
                if (lk := S.is_draft(home)) or code in S.draft_langs:
                    leak(f"/assets/js/{f.name}", "geo banner link", f"{code} {home}", lk or code)
        for tz, cc in (data.get("tz") or {}).items():
            if cc not in (data.get("sites") or {}):
                errors.append(f"/assets/js/{f.name}: time zone {tz} maps to {cc}, which the banner does not offer")


def main():
    here = Path(__file__).resolve().parents[1]
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dist", nargs="?", default="dist", help="the built site (default: ./dist)")
    ap.add_argument("--drafts", action="store_true",
                    help="the dist is a build.py --drafts preview: draft pages must be noindex and marked")
    ap.add_argument("--data", default=str(here / "data"), help="site/data with locales.yaml (default: %(default)s)")
    args = ap.parse_args()
    dist = os.path.abspath(args.dist)
    S = Sections(args.data)
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
        if p.dup_ids:
            errors.append(f"{u}: duplicate element ids {sorted(set(p.dup_ids))}")
        if is_404(u):
            continue
        noindex = u not in indexable
        want = ORIGIN + u
        if not noindex:
            if p.canonical != [want]:
                errors.append(f"{u}: canonical {p.canonical} != [{want}]")
            if p.og_url != want:
                errors.append(f"{u}: og:url {p.og_url} != {want}")
        lang = S.lang(u)
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
        bad = set(codes) - S.live_codes
        if bad:
            errors.append(f"{u}: hreflang codes of no live section {sorted(bad)}")
        amap = dict(p.alternates)
        if "x-default" not in amap:
            errors.append(f"{u}: hreflang cluster without x-default")
        if ORIGIN + u not in amap.values():
            errors.append(f"{u}: hreflang does not reference itself")
        own = S.own[S.of(u)]
        if amap.get(own) != ORIGIN + u:
            errors.append(f"{u}: own code {own} points to {amap.get(own)}")
        for catch, lk in S.catch.items():
            main_code = S.own[lk]
            if main_code in amap and amap.get(catch) != amap[main_code]:
                errors.append(f"{u}: {catch} catch-all {amap.get(catch)} != {main_code} {amap[main_code]}")
            if catch in amap and main_code not in amap:
                errors.append(f"{u}: {catch} catch-all without {main_code}")
        if "en" in amap and "x-default" in amap and amap["en"] != amap["x-default"]:
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
            # every alternate URL's section matches its code
            sec = S.own[S.of(tu)]
            if code != "x-default" and code not in S.catch and code != sec:
                errors.append(f"{u}: hreflang {code} points into section {sec}: {tu}")
        clusters += 1

    # --- sitemaps ---
    ns = {"s": "http://www.sitemaps.org/schemas/sitemap/0.9"}
    in_sitemap = defaultdict(list)
    sitemap_files = []
    try:
        idx = ET.parse(os.path.join(dist, "sitemap.xml")).getroot()
        subs = [e.text for e in idx.findall("s:sitemap/s:loc", ns)]
        if not subs:
            errors.append("sitemap.xml has no child sitemaps")
        for sm in subs:
            sitemap_files.append(sm.rsplit("/", 1)[-1])
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
            want = S.loc[S.of(tu)]["sitemap"]
            if where[0] != want:
                errors.append(f"{loc} is in {where[0]}, expected {want}")
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
        for tag, attr, val, reg, _ in p.links:
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

    # --- draft sections ---
    check_drafts(dist, S, pages, in_sitemap, sitemap_files, args.drafts, errors)

    print(f"pages {len(pages)}, indexable {len(indexable)}, hreflang pages {clusters}, "
          f"sitemap URLs {len(in_sitemap)}, links checked {checked}, JSON-LD blocks {ld_blocks}; "
          f"sections live {len(S.live)}, draft {len(S.draft)}{' (drafts preview)' if args.drafts else ''}")
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
