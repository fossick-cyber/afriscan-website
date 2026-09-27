#!/usr/bin/env python3
"""Crawl a local preview of dist/ over HTTP and check every internal href and src.

    /opt/favhousecheck/.venv/bin/python3 site/tools/crawl.py http://127.0.0.1:5080 [DIST] [--drafts]

python -m http.server does not map /x to x.html the way Cloudflare Pages does (it answers /x with
a folder listing when a folder x/ exists), so extension-less paths are fetched as /x.html.
Stylesheets are fetched and their url(...) references
checked too. Pages in dist/ that the crawl never reaches are reported, and so is every page the crawl
reaches in a section that data/locales.yaml marks status: draft (unless --drafts: a drafts preview).
"""
import argparse
import os
import re
import sys
from pathlib import Path
import urllib.error
import urllib.request
from html.parser import HTMLParser
from urllib.parse import urljoin, urlsplit, unquote

ORIGIN = "https://afri-scan.com"


class Links(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.refs, self.ids = [], set()

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if "id" in a:
            self.ids.add(a["id"])
        if tag == "link" and (a.get("rel") or "") in ("canonical", "alternate"):
            return  # absolute production URLs; checked by check_dist.py
        for k in ("href", "src", "poster"):
            if a.get(k):
                self.refs.append((tag, a[k]))
        for k in ("srcset", "imagesrcset"):
            if a.get(k):
                for part in a[k].split(","):
                    u = part.strip().split(" ")[0]
                    if u:
                        self.refs.append((tag, u))
        if tag == "meta" and a.get("property") == "og:image":
            self.refs.append(("meta", a.get("content")))


SPECIAL = ("404", "thanks", "obrigado", "merci")    # generated 404 and thank-you pages: linked from no page


def section_prefixes(data_dir):
    """{URL prefix: (section, is_draft)} from data/locales.yaml ({} if PyYAML or the file is missing)."""
    try:
        import yaml
        loc = yaml.safe_load(Path(data_dir, "locales.yaml").read_text(encoding="utf-8"))
    except (ImportError, OSError):
        return {}
    return {l["prefix"] + "/": (k, l.get("status", "live") != "live") for k, l in loc.items()}


def main():
    here = Path(__file__).resolve().parents[1]
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("base", help="preview origin, e.g. http://127.0.0.1:5080")
    ap.add_argument("dist", nargs="?", default="dist", help="the directory the preview serves (default: ./dist)")
    ap.add_argument("--drafts", action="store_true", help="the preview is a build.py --drafts build")
    ap.add_argument("--data", default=str(here / "data"), help="site/data with locales.yaml (default: %(default)s)")
    args = ap.parse_args()
    base, dist = args.base.rstrip("/"), args.dist
    sections = section_prefixes(args.data)
    status, cache = {}, {}

    def fetch(path):
        if path in cache:
            return cache[path]
        tried = [path]
        if not path.endswith("/") and not os.path.splitext(path)[1]:
            tried = [path + ".html"]  # Pages serves x.html at /x, even when a folder x/ exists
        res = (404, None, path)
        for p in tried:
            try:
                with urllib.request.urlopen(base + p, timeout=20) as r:
                    body = r.read()
                    res = (r.status, body, p)
                    break
            except urllib.error.HTTPError as e:
                res = (e.code, None, p)
        cache[path] = res
        return res

    errors = []
    queue, seen = ["/"], set()
    ids_by_page, frag_refs = {}, []
    assets = set()
    while queue:
        page = queue.pop()
        if page in seen:
            continue
        seen.add(page)
        code, body, _ = fetch(page)
        if code != 200 or body is None:
            errors.append(f"{page}: HTTP {code}")
            continue
        p = Links()
        p.feed(body.decode("utf-8", "replace"))
        ids_by_page[page] = p.ids
        for tag, ref in p.refs:
            if ref.startswith(("mailto:", "tel:", "data:", "javascript:")):
                continue
            if ref.startswith(ORIGIN):
                ref = ref[len(ORIGIN):] or "/"
            parts = urlsplit(urljoin(page, ref))
            if parts.scheme in ("http", "https") and parts.netloc and not parts.netloc.startswith("127.0.0.1"):
                continue
            path = parts.path or page
            if parts.fragment:
                frag_refs.append((page, path, unquote(parts.fragment)))
            if tag == "a" and (path.endswith("/") or not os.path.splitext(path)[1]):
                queue.append(path)
            else:
                assets.add((page, path))
    checked_assets = 0
    css_seen = set()
    for page, path in sorted(assets):
        code, body, _ = fetch(path)
        checked_assets += 1
        if code != 200:
            errors.append(f"{page}: asset {path} HTTP {code}")
        elif path.endswith(".css") and path not in css_seen:
            css_seen.add(path)
            for u in re.findall(r"url\(\s*['\"]?([^'\")]+)", body.decode("utf-8", "replace")):
                if u.startswith("data:"):
                    continue
                cp = urlsplit(urljoin(path, u)).path
                c2, _, _ = fetch(cp)
                checked_assets += 1
                if c2 != 200:
                    errors.append(f"{path}: url({u}) HTTP {c2}")
    for page, path, frag in frag_refs:
        if path not in ids_by_page:
            continue
        if frag not in ids_by_page[path]:
            errors.append(f"{page}: #{frag} missing on {path}")

    # pages in dist that the crawl never reached
    built = set()
    for root, _, files in os.walk(dist):
        for fn in files:
            if fn.endswith(".html"):
                rel = os.path.relpath(os.path.join(root, fn), dist).replace(os.sep, "/")
                if rel == "index.html":
                    built.add("/")
                elif rel.endswith("index.html"):
                    built.add("/" + rel[:-10])
                else:
                    built.add("/" + rel[:-5])
    unreached = sorted(u for u in built - seen if u.rstrip("/").rsplit("/", 1)[-1] not in SPECIAL)
    for u in unreached:
        errors.append(f"{u}: built but not reachable by crawling from /")
    for u in sorted(seen) if not args.drafts else []:
        pre = max((p for p in sections if u.startswith(p)), key=len, default=None)
        if pre and sections[pre][1]:
            errors.append(f"{u}: reached by crawling, but it is in the draft section {sections[pre][0]}")
    print(f"crawled {len(seen)} pages, checked {checked_assets} asset refs and {len(frag_refs)} anchors")
    for e in errors:
        print("ERROR", e)
    print(f"{len(errors)} errors")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
