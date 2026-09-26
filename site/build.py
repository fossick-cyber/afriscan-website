#!/usr/bin/env python3
"""Build afri-scan.com (global, /mz/, /mz/pt/, /za/, /ng/) from site/ into dist/.

    /opt/favhousecheck/.venv/bin/python3 site/build.py            # build + guards, exit 1 on any error
    /opt/favhousecheck/.venv/bin/python3 site/build.py --drafts   # also build status: draft pages (local preview only)
    /opt/favhousecheck/.venv/bin/python3 site/build.py --selftest # prove the guards fail on seeded mistakes

Sources: site/content (pages), site/data (facts, UI strings, rules), site/templates, site/static,
site/images. Everything the build writes goes to dist/, which is what Cloudflare Pages serves.
See CONTENT_GUIDE.md for how to add pages.
"""
import argparse
import datetime as dt
import hashlib
import html as htmllib
import json
import re
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import yaml
from jinja2 import Environment, FileSystemLoader, StrictUndefined
from markupsafe import Markup, escape

SITE = Path(__file__).resolve().parent
ROOT = SITE.parent
sys.path.insert(0, str(SITE))

from lib import brand  # noqa: E402
from lib.content import (Block, ContentError, add_heading_ids, make_markdown, parse_blocks,  # noqa: E402
                         slugify, split_front_matter, wrap_tables)
from lib.images import ImageError, Images  # noqa: E402

TEMPLATES = {"home", "page", "industry", "solution", "country_home", "law", "law_hub", "article", "hub", "contact"}
DEFAULT_SECTION = {"industry": "industries", "solution": "solutions", "article": "insights"}
HUBS = {"industries", "solutions", "how", "countries", "resources", "insights"}
NAV_GROUPS = ["industries", "solutions", "how", "countries", "resources"]
TONES = {"light", "alt", "dark", "brand"}

# Components usable in content bodies:  :::name{attr="value"} ... :::
COMPONENTS = {
    "section": {"allowed": {"id", "tone", "title", "eyebrow", "lead", "width", "class"}},
    "cards": {"allowed": {"cols", "style"}},
    "card": {"required": {"title"}, "allowed": {"href", "icon", "eyebrow", "tag", "key", "cta"}},
    "steps": {"allowed": {"style"}},
    "step": {"required": {"title"}, "allowed": {"icon"}},
    "callout": {"allowed": {"tone", "title", "icon"}},
    "columns": {"allowed": {"split", "align"}},
    "col": {"allowed": {"class"}},
    "figure": {"required": {"src", "alt"}, "allowed": {"caption", "credit", "badge", "size", "priority", "sizes"}},
    "facts": {"allowed": {"cols"}},
    "chips": {"allowed": set()},
    "checklist": {"allowed": {"tone"}},
    "cta": {"allowed": {"title", "text", "button", "href", "intent", "secondary", "secondary_href"}},
    "catalogue": {"allowed": {"groups", "services", "style"}},
    "solutions": {"allowed": {"keys", "group", "cols"}},
    "industries": {"allowed": {"keys", "tier", "cols"}},
    "pages": {"required": {"section"}, "allowed": {"limit", "cols"}},
    "sources": {"required": {"law"}, "allowed": {"ids"}},
    "law-table": {"allowed": set()},
    "segments": {"required": {"data"}, "allowed": set()},
    "register": {"required": {"data"}, "allowed": {"limit"}},
    "details": {"required": {"summary"}, "allowed": {"open"}},
    "lead": {"allowed": set()},
    "country-sites": {"allowed": {"match"}},
}


def load_yaml(p):
    return yaml.safe_load(Path(p).read_text(encoding="utf-8"))


def sha(data, n=8):
    return hashlib.sha256(data if isinstance(data, bytes) else data.encode()).hexdigest()[:n]


class Build:
    def __init__(self, content_dir=SITE / "content", dist=ROOT / "dist", drafts=False, quiet=False,
                 today=None, law_dir=SITE / "data/law"):
        self.content_dir, self.dist, self.drafts, self.quiet = content_dir, dist, drafts, quiet
        self.today = today or dt.date.today()
        self.errors, self.warnings = [], []
        self.site = load_yaml(SITE / "data/site.yaml")
        self.base = self.site["base_url"].rstrip("/")
        self.locales = load_yaml(SITE / "data/locales.yaml")
        self.i18n = {n: load_yaml(SITE / f"data/i18n/{n}.yaml") for n in {l["i18n"] for l in self.locales.values()}}
        self.catalogue = load_yaml(SITE / "data/catalogue.yaml")
        self.rules = load_yaml(SITE / "data/rules.yaml")
        self.glossary = load_yaml(SITE / "data/glossary/pt-MZ.yaml")
        self.redirects = load_yaml(SITE / "data/redirects.yaml")
        self.icons = load_yaml(SITE / "data/icons.yaml")
        self.law = {p.stem: load_yaml(p) for p in sorted(Path(law_dir).glob("*.yaml"))}
        self.samples = {p.stem: json.loads(p.read_text(encoding="utf-8"))
                        for p in sorted((SITE / "data/samples").glob("*.json"))}
        self.md = make_markdown()
        self.cat_ind = {i["key"]: i for i in self.catalogue["industries"]}
        self.cat_sol = {s["key"]: s for s in self.catalogue["solutions"]}
        self.cat_srv = {s["id"]: s for s in self.catalogue["services"]}

    # ------------------------------------------------------------------ messages
    def err(self, msg):
        self.errors.append(msg)

    def warn(self, msg):
        self.warnings.append(msg)

    # ------------------------------------------------------------------ pages
    def read_pages(self):
        pages = []
        for lk, loc in self.locales.items():
            base = self.content_dir / loc["content"]
            if not base.exists():
                continue
            for f in sorted(base.rglob("*.md")):
                rel = f.relative_to(base)
                if loc["content"] == "mz/en" and rel.parts and rel.parts[0] == "pt":
                    continue
                try:
                    p = self.read_page(f, lk, rel)
                except ContentError as e:
                    self.err(str(e))
                    continue
                if p:
                    pages.append(p)
        return pages

    def read_page(self, f, lk, rel):
        loc = self.locales[lk]
        where = str(f.relative_to(ROOT)) if f.is_relative_to(ROOT) else str(f)
        meta, body = split_front_matter(f.read_text(encoding="utf-8"), where)
        for req in ("title", "description", "h1"):
            if not meta.get(req):
                raise ContentError(f"{where}: front matter needs '{req}'")
        stem = rel.with_suffix("").as_posix()
        if stem.endswith("/index"):
            raise ContentError(f"{where}: only the section home may be called index.md")
        slug = meta.get("slug", "" if stem == "index" else stem)
        if slug != slug.lower() or slug.endswith("/") or not re.fullmatch(r"[a-z0-9]+(?:[-/][a-z0-9]+)*|", slug):
            raise ContentError(f"{where}: slug {slug!r} must be lowercase ASCII words joined by - or /, no trailing slash")
        if lk == "mz-en" and (slug == "pt" or slug.startswith("pt/")):
            raise ContentError(f"{where}: slug 'pt' is reserved under /mz/ for the Portuguese section")
        if slug.split("/")[0] in {"assets", "geo", "404", "thanks", "obrigado"}:
            raise ContentError(f"{where}: slug {slug!r} is reserved")
        template = meta.get("template", "home" if not slug else "page")
        if template not in TEMPLATES:
            raise ContentError(f"{where}: unknown template {template!r} (use one of {sorted(TEMPLATES)})")
        status = meta.get("status", "published")
        if status not in ("published", "draft"):
            raise ContentError(f"{where}: status must be 'published' or 'draft'")
        if status == "draft" and not self.drafts:
            return None
        prefix = loc["prefix"]
        url = f"{prefix}/{slug}" if slug else f"{prefix}/"
        out = self.dist / prefix.lstrip("/") / (f"{slug}.html" if slug else "index.html")
        key = meta.get("key") or ("home" if not slug else f"{lk}:{slug}")
        hub = meta.get("hub")
        if hub and hub not in HUBS:
            raise ContentError(f"{where}: hub must be one of {sorted(HUBS)}")
        section = meta.get("section", DEFAULT_SECTION.get(template))
        if template == "law" and not meta.get("law"):
            raise ContentError(f"{where}: a law page needs 'law: <country code>' (data/law/<cc>.yaml)")
        if template == "law" and meta["law"] not in self.law:
            raise ContentError(f"{where}: no data/law/{meta['law']}.yaml for this law page")
        if template == "article" and not meta.get("published"):
            raise ContentError(f"{where}: an article needs 'published: YYYY-MM-DD'")
        if lk == "mz-pt" and not meta.get("reviewed_on"):
            self.warn(f"{url}: Portuguese page awaiting native Mozambican review (no reviewed_on)")
        return dict(meta=meta, body=body, src=f, where=where, loc_key=lk, loc=loc, t=self.i18n[loc["i18n"]],
                    slug=slug, url=url, abs_url=self.base + url, out=out, key=key, template=template,
                    status=status, hub=hub, section=section, title=meta["title"].strip(),
                    description=" ".join(str(meta["description"]).split()), h1=meta["h1"].strip(),
                    lang=loc["lang"], lang2=loc["lang"][:2], noindex=bool(meta.get("noindex")) or status != "published")

    # ------------------------------------------------------------------ lookup helpers
    def find(self, key, lk):
        return self.by_key.get(key, {}).get(lk)

    def resolve(self, key, lk):
        """Same-section page for a cluster key, else the global one (English sections only get global)."""
        return self.find(key, lk) or self.find(key, "global")

    def home_of(self, lk):
        return self.find("home", lk)

    def label_of(self, p):
        m = p["meta"]
        if m.get("crumb"):
            return m["crumb"]
        if m.get("nav_label"):
            return m["nav_label"]
        lang = "pt" if p["lang2"] == "pt" else "en"
        cat = self.cat_ind.get(p["key"]) or self.cat_sol.get(p["key"])
        if cat:
            return cat["name"][lang]
        return p["h1"]

    def hub_page(self, group, lk):
        return self.hubs.get((group, lk)) or (self.hubs.get((group, "global")) if lk != "global" else None)

    def contact_url(self, p, intent="proposal", extra=None):
        c = self.resolve("contact", p["loc_key"])
        href = c["url"] if c else "/contact"
        q = {"intent": intent} if intent else {}
        if p["key"] in self.cat_ind:
            q["industry"] = p["key"]
        if p["key"] in self.cat_sol:
            q["service"] = p["key"]
        if p["loc"]["country"]:
            q["country"] = p["loc"]["country"].lower()
        q.update(extra or {})
        return href + ("?" + "&".join(f"{k}={v}" for k, v in q.items()) if q else "")

    def fmt_date(self, d, t):
        if isinstance(d, str):
            d = dt.date.fromisoformat(d)
        return t["date_format"].format(d=d.day, month=t["months"][d.month - 1], y=d.year)

    # ------------------------------------------------------------------ structure
    def index(self, pages):
        self.pages = pages
        self.by_url = {}
        for p in pages:
            if p["url"] in self.by_url:
                self.err(f"duplicate URL {p['url']}: {p['where']} and {self.by_url[p['url']]['where']}")
            self.by_url[p["url"]] = p
        self.by_key = defaultdict(dict)
        for p in pages:
            if p["loc_key"] in self.by_key[p["key"]]:
                other = self.by_key[p["key"]][p["loc_key"]]
                self.err(f"cluster '{p['key']}' has two {p['loc_key']} pages: {p['where']} and {other['where']}")
            self.by_key[p["key"]][p["loc_key"]] = p
        self.hubs = {}
        for p in pages:
            if p["hub"]:
                if (p["hub"], p["loc_key"]) in self.hubs:
                    self.err(f"two '{p['hub']}' hubs in {p['loc_key']}")
                self.hubs[(p["hub"], p["loc_key"])] = p
        self.live_locales = [lk for lk in self.locales if self.home_of(lk)]
        for p in pages:
            p["alternates"] = self.alternates(p)
            p["region_links"] = self.region_links(p)
            p["crumbs"] = self.breadcrumbs(p)
        for lk in self.locales:
            if lk != "global" and not self.home_of(lk) and any(p["loc_key"] == lk for p in pages):
                self.err(f"section {lk} has pages but no home page (content/{self.locales[lk]['content']}/index.md)")

    def alternates(self, p):
        if p["noindex"]:
            return []
        cluster = [m for m in self.by_key[p["key"]].values() if not m["noindex"]]
        if len(cluster) < 2:
            return []
        alts, taken = [], set()
        order = list(self.locales)
        for m in sorted(cluster, key=lambda m: order.index(m["loc_key"]) if m["loc_key"] != "global" else -1):
            for code in self.locales[m["loc_key"]]["hreflang"]:
                if code not in taken:
                    alts.append({"code": code, "href": m["abs_url"]})
                    taken.add(code)
        if "x-default" not in taken:
            en = [m for m in cluster if m["lang2"] == "en"]
            if en:
                alts.append({"code": "x-default", "href": en[0]["abs_url"]})
        return alts

    def region_links(self, p):
        links = []
        for lk in self.live_locales:
            target = self.find(p["key"], lk) or self.home_of(lk)
            loc = self.locales[lk]
            links.append({"key": lk, "label": loc["label"], "lang": loc["lang"], "href": target["url"],
                          "current": lk == p["loc_key"]})
        return links

    def breadcrumbs(self, p):
        if p["template"] == "home":
            return []
        t = p["t"]
        crumbs = [{"name": t["home_crumb"], "url": "/"}]
        lk = p["loc_key"]
        home = self.home_of(lk)
        if p["template"] == "country_home":
            return crumbs + [{"name": p["loc"]["country_name"], "url": p["url"]}]
        if lk != "global" and home:
            crumbs.append({"name": p["loc"]["country_name"], "url": home["url"]})
        sec = p["hub"] and None
        section = p["section"]
        if section:
            hub = self.hubs.get((section, lk)) or (self.hubs.get((section, "global")) if lk == "global" else None)
            if hub and hub is not p:
                crumbs.append({"name": self.label_of(hub), "url": hub["url"]})
        parent = p["meta"].get("parent")
        if parent:
            pp = self.resolve(parent, lk)
            if not pp:
                self.err(f"{p['url']}: parent '{parent}' is not a built page key")
            elif pp["url"] not in {c["url"] for c in crumbs}:
                crumbs.append({"name": self.label_of(pp), "url": pp["url"]})
        crumbs.append({"name": self.label_of(p), "url": p["url"]})
        return crumbs

    # ------------------------------------------------------------------ navigation
    def nav_for(self, lk):
        lang = "pt" if self.locales[lk]["lang"].startswith("pt") else "en"
        t = self.i18n[self.locales[lk]["i18n"]]
        page_lang = self.locales[lk]["lang"]

        def item(p, label=None, blurb=None, **kw):
            d = {"label": label or self.label_of(p), "href": p["url"],
                 "blurb": blurb if blurb is not None else p["meta"].get("nav_blurb", ""),
                 "lang": p["lang"] if p["lang2"] != page_lang[:2] else None}
            d.update(kw)
            return d

        groups = []
        # industries / solutions from the catalogue
        for gname, cat in (("industries", self.catalogue["industries"]), ("solutions", self.catalogue["solutions"])):
            items = []
            for c in cat:
                if lk in c.get("exclude", []):
                    continue
                p = self.resolve(c["key"], lk)
                if p and not p["noindex"]:
                    items.append(item(p, c["name"][lang], (c.get("menu") or c["blurb"])[lang],
                                      tier=c.get("tier", 1), group=c.get("group"), icon=c.get("icon"),
                                      key=c["key"]))
            for p in self.pages:              # country-only pages that join the group
                if (p["loc_key"] == lk and p["meta"].get("nav_group") == gname and p["key"] not in self.cat_ind
                        and p["key"] not in self.cat_sol and not p["noindex"]):
                    items.append(item(p, tier=2, group=p["meta"].get("nav_subgroup")))
            hub = self.hub_page(gname, lk)
            if items or hub:
                groups.append({"key": gname, "label": t["nav"][gname], "href": (hub or {"url": items[0]["href"]})["url"],
                               "items": items})
        # how / resources: pages that set nav_group (this section first, then global by key)
        for gname in ("how", "countries", "resources"):
            items, seen = [], set()
            cands = [p for p in self.pages if p["meta"].get("nav_group") == gname and not p["noindex"]
                     and lang in (p["meta"].get("nav_langs") or [lang])]
            cands.sort(key=lambda p: (p["meta"].get("nav_order", 50), p["title"]))
            for p in cands:
                if p["loc_key"] == lk or (p["loc_key"] == "global" and not self.find(p["key"], lk)):
                    if p["key"] in seen:
                        continue
                    seen.add(p["key"])
                    items.append(item(p))
            if gname == "countries":
                homes = [item(self.home_of(k), self.locales[k]["label"], "", lang=None if self.locales[k]["lang"][:2] == page_lang[:2] else self.locales[k]["lang"])
                         for k in self.live_locales if k != "global"]
                items = homes + items
            hub = self.hub_page(gname, lk)
            if items or hub:
                groups.append({"key": gname, "label": t["nav"][gname],
                               "href": (hub or {"url": items[0]["href"]})["url"], "items": items})
        order = {g: i for i, g in enumerate(NAV_GROUPS)}
        groups.sort(key=lambda g: order[g["key"]])
        return groups

    # ------------------------------------------------------------------ rendering: components
    def icon(self, name, cls="icon"):
        if name not in self.icons:
            self.err(f"unknown icon '{name}' (see site/data/icons.yaml)")
            return Markup("")
        return Markup(f'<svg class="{cls}" viewBox="0 0 24 24" aria-hidden="true" focusable="false">'
                      f'{self.icons[name]}</svg>')

    def md_inline(self, text):
        return Markup(self.md.renderInline(str(text or "")))

    def md_block(self, text):
        return Markup(self.md.render(str(text or "")))

    def render_body(self, p, body):
        nodes = parse_blocks(body, p["where"])
        used_ids = set()
        out = []
        pending = []

        def flush_default():
            if pending:
                inner = self.render_nodes(p, pending)
                if inner.strip():
                    out.append(f'<section class="section section--light"><div class="container flow">{inner}</div></section>')
                pending.clear()

        for n in nodes:
            if isinstance(n, Block) and n.name == "section":
                flush_default()
                out.append(self.render_block(p, n))
            else:
                pending.append(n)
        flush_default()
        html = "\n".join(out)
        html = add_heading_ids(html, used_ids)
        html = wrap_tables(html, "Tabela" if p["lang2"] == "pt" else "Table")
        html = self.resolve_key_links(p, html)
        return html

    def render_nodes(self, p, nodes):
        parts, holes = [], {}
        for i, n in enumerate(nodes):
            if isinstance(n, str):
                parts.append(n)
            else:
                token = f"<!--cmp-{id(n)}-{i}-->"
                holes[token] = self.render_block(p, n)
                parts.append(f"\n\n{token}\n\n")
        html = self.md.render("".join(parts))
        for token, h in holes.items():
            html = html.replace(token, h)
        return html

    def render_block(self, p, b: Block):
        spec = COMPONENTS.get(b.name)
        where = f"{p['where']}:{b.line}"
        if not spec:
            raise ContentError(f"{where}: unknown component ':::{b.name}' (see CONTENT_GUIDE.md)")
        allowed = spec.get("allowed", set()) | spec.get("required", set())
        for a in b.attrs:
            if a not in allowed:
                raise ContentError(f"{where}: ':::{b.name}' has no attribute '{a}' (allowed: {sorted(allowed)})")
        for a in spec.get("required", set()):
            if not b.attrs.get(a):
                raise ContentError(f"{where}: ':::{b.name}' needs {a}=\"…\"")
        raw = "".join(c for c in b.children if isinstance(c, str))
        inner = Markup(self.render_nodes(p, b.children))
        ctx = self.page_ctx(p)
        ctx.update(a=b.attrs, inner=inner, raw=raw, where=where, block=b)
        method = getattr(self, "cmp_" + b.name.replace("-", "_"), None)
        if method:
            ctx.update(method(p, b, ctx) or {})
        try:
            return self.env.get_template(f"components/{b.name}.html.j2").render(**ctx).strip()
        except ContentError:
            raise
        except Exception as e:
            raise ContentError(f"{where}: ':::{b.name}' failed to render: {e}")

    # component data helpers (return extra template context)
    def cmp_section(self, p, b, ctx):
        tone = b.attrs.get("tone", "light")
        if tone not in TONES:
            raise ContentError(f"{ctx['where']}: tone must be one of {sorted(TONES)}")
        return {"tone": tone}

    def cmp_card(self, p, b, ctx):
        href = b.attrs.get("href")
        if b.attrs.get("key"):
            key, _, frag = b.attrs["key"].partition("#")
            tp = self.resolve(key, p["loc_key"])
            if not tp:
                if key in self.cat_ind or key in self.cat_sol:
                    return {"href": None}      # catalogue item without a page yet: a plain card, as on hubs
                raise ContentError(f"{ctx['where']}: card key '{key}' is not a built page")
            href = tp["url"]
            if frag:
                # pages render in content order, so a target that is already rendered can be checked
                # here; an anchor it lacks is dropped (card links to the page top) with a warning.
                if tp.get("html") and f'id="{frag}"' not in tp["html"]:
                    self.warn(f"{p['url']}: card key '{b.attrs['key']}': no id '{frag}' on {tp['url']}; linking to the page")
                else:
                    href += "#" + frag
        return {"href": href}

    def cmp_facts(self, p, b, ctx):
        rows = []
        for line in ctx["raw"].splitlines():
            m = re.match(r"^\s*[-*]\s+(.+?):\s+(.+)$", line)
            if m:
                rows.append((self.md_inline(m.group(1).strip("* ")), self.md_inline(m.group(2))))
            elif line.strip():
                raise ContentError(f"{ctx['where']}: facts lines look like '- Term: value' (got {line.strip()!r})")
        return {"rows": rows}

    def _cards_from_catalogue(self, p, items):
        lang = "pt" if p["lang2"] == "pt" else "en"
        cards = []
        for c in items:
            if p["loc_key"] in c.get("exclude", []):
                continue
            tp = self.resolve(c["key"], p["loc_key"])
            cards.append({"title": c["name"][lang], "text": c["blurb"][lang], "icon": c.get("icon"),
                          "href": tp["url"] if tp and not tp["noindex"] else None, "code": c.get("code"),
                          "services": c.get("services", []), "flagship": c.get("flagship")})
        return cards

    def cmp_industries(self, p, b, ctx):
        items = self.catalogue["industries"]
        if b.attrs.get("keys"):
            keys = [k.strip() for k in b.attrs["keys"].split(",")]
            bad = [k for k in keys if k not in self.cat_ind]
            if bad:
                raise ContentError(f"{ctx['where']}: unknown industry keys {bad}")
            items = [self.cat_ind[k] for k in keys]
        if b.attrs.get("tier"):
            items = [i for i in items if str(i.get("tier")) == b.attrs["tier"]]
        return {"cards": self._cards_from_catalogue(p, items)}

    def cmp_solutions(self, p, b, ctx):
        items = self.catalogue["solutions"]
        if b.attrs.get("keys"):
            keys = [k.strip() for k in b.attrs["keys"].split(",")]
            bad = [k for k in keys if k not in self.cat_sol]
            if bad:
                raise ContentError(f"{ctx['where']}: unknown solution keys {bad}")
            items = [self.cat_sol[k] for k in keys]
        if b.attrs.get("group"):
            items = [s for s in items if s["group"] == b.attrs["group"]]
        return {"cards": self._cards_from_catalogue(p, items)}

    def services_for(self, p, ids=None, groups=None):
        lang = "pt" if p["lang2"] == "pt" else "en"
        out = []
        for s in self.catalogue["services"]:
            if ids and s["id"] not in ids:
                continue
            if groups and s["group"] not in groups:
                continue
            if s.get("only") and p["loc_key"] not in s["only"]:
                continue
            if p["loc_key"] in s.get("exclude", []):
                continue
            name = s.get("name_pt") if lang == "pt" and s.get("name_pt") else s["name"]
            line = s.get("line_pt") if lang == "pt" and s.get("line_pt") else s["line"]
            out.append({"id": s["id"], "name": name, "line": line, "group": s["group"]})
        return out

    def cmp_catalogue(self, p, b, ctx):
        lang = "pt" if p["lang2"] == "pt" else "en"
        ids = [x.strip() for x in b.attrs["services"].split(",")] if b.attrs.get("services") else None
        groups = [x.strip() for x in b.attrs["groups"].split(",")] if b.attrs.get("groups") else None
        if ids:
            bad = [i for i in ids if i not in self.cat_srv]
            if bad:
                raise ContentError(f"{ctx['where']}: unknown service ids {bad}")
        services = self.services_for(p, ids, groups)
        by_group = []
        for g in self.catalogue["service_groups"]:
            rows = [s for s in services if s["group"] == g["key"]]
            if rows:
                by_group.append({"key": g["key"], "name": g["name"][lang], "services": rows})
        return {"groups": by_group, "services": services}

    def cmp_pages(self, p, b, ctx):
        sec = b.attrs["section"]
        lk = p["loc_key"]
        items = [q for q in self.pages if q["section"] == sec and q is not p and not q["noindex"]
                 and (q["loc_key"] == lk or (q["loc_key"] == "global" and not self.find(q["key"], lk)
                                             and p["lang2"] == "en"))]
        items.sort(key=lambda q: (q["meta"].get("nav_order", 50), str(q["meta"].get("published", "")), q["title"]))
        if sec == "insights":
            items.sort(key=lambda q: str(q["meta"].get("published", "")), reverse=True)
        limit = int(b.attrs.get("limit", 0) or 0)
        if limit:
            items = items[:limit]
        cards = [{"title": self.label_of(q), "text": q["meta"].get("summary") or q["description"], "href": q["url"],
                  "icon": q["meta"].get("icon"), "date": q["meta"].get("published")} for q in items]
        return {"cards": cards}

    def law_for(self, p, cc):
        """data/law/<cc>.yaml with each instrument's optional title_pt / identifier_pt / note_pt used on PT pages."""
        law = self.law.get(cc)
        if not law or p["lang2"] != "pt":
            return law
        loc = lambda i: {**i, **{k: i[k + "_pt"] for k in ("title", "identifier", "note") if i.get(k + "_pt")}}
        return {**law, "instruments": [loc(i) for i in law.get("instruments", [])]}

    def cmp_sources(self, p, b, ctx):
        law = self.law_for(p, b.attrs["law"])
        if not law:
            raise ContentError(f"{ctx['where']}: no data/law/{b.attrs['law']}.yaml")
        ids = [x.strip() for x in b.attrs.get("ids", "").split(",") if x.strip()]
        inst = law["instruments"]
        if ids:
            known = {i["id"] for i in inst}
            bad = [i for i in ids if i not in known]
            if bad:
                raise ContentError(f"{ctx['where']}: unknown instrument ids {bad}")
            inst = [i for i in inst if i["id"] in ids]
        return {"instruments": inst}

    def cmp_law_table(self, p, b, ctx):
        rows = []
        for cc, law in self.law.items():
            guide = self.resolve(f"{cc}-drone-law", p["loc_key"]) or next(
                (q for q in self.pages if q["template"] == "law" and q["meta"].get("law") == cc
                 and q["lang2"] == p["lang2"] and not q["noindex"]), None)
            if not guide:
                continue
            rows.append({"cc": cc, "law": law, "guide": guide})
        return {"rows": rows}

    def cmp_country_sites(self, p, b, ctx):
        # match="page": link each country to its version of this page (same cluster key), else its home
        match = b.attrs.get("match") == "page"
        sites = []
        for k in self.live_locales:
            if k == "global":
                continue
            target = self.find(p["key"], k) if match else None
            if not target or target["noindex"]:
                target = self.home_of(k)
            sites.append({"label": self.locales[k]["label"], "href": target["url"], "lang": self.locales[k]["lang"]})
        return {"sites": sites}

    def cmp_segments(self, p, b, ctx):
        data = self.samples.get(b.attrs["data"])
        if not data:
            raise ContentError(f"{ctx['where']}: no data/samples/{b.attrs['data']}.json")
        segs = data["segments"]
        total = segs[-1]["to_m"]
        return {"data": data, "segs": segs, "total": total, "maxc": max(s["count"] for s in segs) or 1}

    def cmp_register(self, p, b, ctx):
        data = self.samples.get(b.attrs["data"])
        if not data:
            raise ContentError(f"{ctx['where']}: no data/samples/{b.attrs['data']}.json")
        rows = data["register_excerpt"]
        if b.attrs.get("limit"):
            rows = rows[: int(b.attrs["limit"])]
        return {"data": data, "rows": rows}

    def cmp_figure(self, p, b, ctx):
        size = b.attrs.get("size", "wide")
        sizes = b.attrs.get("sizes") or {"full": "100vw", "wide": "(max-width: 1188px) 100vw, 1140px",
                                          "narrow": "(max-width: 828px) 100vw, 780px",
                                          "half": "(max-width: 900px) 100vw, 560px",
                                          "third": "(max-width: 640px) 100vw, (max-width: 1000px) 50vw, 370px"}.get(size, "100vw")
        try:
            pic = self.images.picture(b.attrs["src"], b.attrs["alt"], sizes=sizes,
                                      priority=b.attrs.get("priority") == "true", page_url=p["url"])
        except ImageError as e:
            raise ContentError(f"{ctx['where']}: {e}")
        return {"picture": Markup(pic), "size": size}

    def resolve_key_links(self, p, html):
        def sub(m):
            key, frag = m.group(1), m.group(2) or ""
            tp = self.resolve(key, p["loc_key"])
            if not tp:
                self.err(f"{p['url']}: link to key:{key} but no page has that key")
                return f'href="#missing-{key}"'
            return f'href="{tp["url"]}{frag}"'

        return re.sub(r'href="key:([a-z0-9:/-]+)(#[\w-]+)?"', sub, html)

    # ------------------------------------------------------------------ rendering: pages
    def page_ctx(self, p):
        lk = p["loc_key"]
        return dict(page=p, loc=p["loc"], t=p["t"], site=self.site, nav=self.navs[lk], build=self,
                    icon=self.icon, md=self.md_inline, mdblock=self.md_block, contact=self.contact_url,
                    fmt_date=lambda d: self.fmt_date(d, p["t"]), assets=self.asset_urls,
                    lang_key="pt" if p["lang2"] == "pt" else "en", catalogue=self.catalogue,
                    resolve=lambda k: self.resolve(k, lk), footer=self.footers[lk], today=self.today,
                    region_js=self.region_js_url, thanks_url=self.thanks_url(p), label_of=self.label_of)

    def thanks_url(self, p):
        if p["lang2"] == "pt" and self.home_of("mz-pt"):
            return f"{self.base}/mz/pt/obrigado"
        return f"{self.base}/thanks"

    def footer_for(self, lk):
        nav = self.navs[lk]
        g = {x["key"]: x for x in nav}
        explore = [{"label": x["label"], "href": x["href"]} for x in nav]
        return {"explore": explore,
                "industries": g.get("industries", {}).get("items", []),
                "solutions": self.footer_solutions(g.get("solutions", {}).get("items", [])),
                "how": g.get("how", {}).get("items", []),
                "resources": g.get("resources", {}).get("items", []),
                "countries": [{"label": self.locales[k]["label"], "href": self.home_of(k)["url"],
                               "lang": self.locales[k]["lang"]} for k in self.live_locales]}

    def footer_solutions(self, items):
        """Solutions marked `footer: true` in catalogue.yaml, in catalogue order; else the first eight."""
        flagged = {s["key"] for s in self.catalogue["solutions"] if s.get("footer")}
        return [i for i in items if i.get("key") in flagged] if flagged else items[:8]

    def related_cards(self, p, keys):
        lang = "pt" if p["lang2"] == "pt" else "en"
        cards = []
        for k in keys or []:
            tp = self.resolve(k, p["loc_key"])
            cat = self.cat_ind.get(k) or self.cat_sol.get(k)
            if cat and p["loc_key"] in cat.get("exclude", []):
                continue
            if not tp:
                if cat:
                    continue          # catalogue item without a page yet: skip quietly
                self.err(f"{p['url']}: related key '{k}' is not a built page")
                continue
            cards.append({"title": cat["name"][lang] if cat else self.label_of(tp),
                          "text": cat["blurb"][lang] if cat else tp["description"],
                          "href": tp["url"], "icon": (cat or {}).get("icon") or tp["meta"].get("icon")})
        return cards

    def used_in(self, p):
        lang = "pt" if p["lang2"] == "pt" else "en"
        out = []
        for k in p["meta"].get("used_in", []) or []:
            if k not in self.cat_ind:
                self.err(f"{p['url']}: used_in '{k}' is not an industry key")
                continue
            tp = self.resolve(k, p["loc_key"])
            out.append({"label": self.cat_ind[k]["name"][lang], "href": tp["url"] if tp else None})
        return out

    def render_page(self, p):
        m = p["meta"]
        try:
            if p["template"] == "contact":
                html = self.render_nodes(p, parse_blocks(p["body"], p["where"]))
                p["body_html"] = Markup(self.resolve_key_links(p, html))
            else:
                p["body_html"] = Markup(self.render_body(p, p["body"]))
        except (ContentError, ImageError) as e:
            self.err(str(e))
            p["body_html"] = Markup("")
        p["faq"] = [{"q": f["q"], "a": self.md_block(f["a"])} for f in m.get("faq", []) or []]
        for f in m.get("faq", []) or []:
            if not f.get("q") or not f.get("a"):
                self.err(f"{p['url']}: every faq item needs q and a")
        p["related"] = self.related_cards(p, m.get("related"))
        tones = re.findall(r'<section class="section section--(\w+)', str(p["body_html"]))
        last = tones[-1] if tones else ("dark" if p["template"] in ("home", "country_home") else "light")
        flip = lambda t: "light" if t == "alt" else "alt"
        p["related_tone"] = flip(last)
        p["faq_tone"] = flip(p["related_tone"] if m.get("related") else last)
        p["used"] = self.used_in(p)
        p["lead_html"] = self.md_inline(m.get("lead", "")) if m.get("lead") else ""
        hero = m.get("hero") or {}
        p["hero_picture"] = None
        if hero.get("image"):
            try:
                p["hero_picture"] = Markup(self.images.picture(
                    hero["image"], hero.get("alt", ""), sizes="100vw", priority=True, cls="hero-img",
                    page_url=p["url"]))
            except ImageError as e:
                self.err(f"{p['url']}: {e}")
        p["cta"] = self.cta_for(p)
        p["og_image"] = self.og_for(p)
        p["jsonld"] = self.jsonld(p)
        ctx = self.page_ctx(p)
        if p["template"] == "law":
            ctx["law"] = self.law_for(p, m["law"])
        if p["template"] in ("law", "law_hub"):
            ctx["as_of"] = m.get("as_of", self.site["law_as_of"])
        html = self.env.get_template(f"{p['template']}.html.j2").render(**ctx)
        html = re.sub(r"\n\s*\n+", "\n", html)
        p["html"] = html
        p["out"].parent.mkdir(parents=True, exist_ok=True)
        p["out"].write_text(html, encoding="utf-8")

    def cta_for(self, p):
        c = p["meta"].get("cta", {})
        if c is False:
            return None
        t = p["t"]["cta"]
        c = c or {}
        return {"title": c.get("title", t["title"]), "text": c.get("text", t["text"]),
                "button": c.get("button", t["button"]),
                "href": c.get("href") or self.contact_url(p, c.get("intent", "proposal")),
                "secondary": c.get("secondary"), "secondary_href": c.get("secondary_href")}

    def og_for(self, p):
        og = p["meta"].get("og") or {}
        region = p["loc"]["country_name"] or ""
        head = og.get("headline") or p["h1"]
        sub = og.get("subline") or p["t"]["footer"]["review"]
        name = brand.card(self.dist / "assets/og", p["url"], head, sub, region, self.font, CACHE_DIR / "og")
        return {"url": f"{self.base}/assets/og/{name}", "alt": og.get("alt") or head}

    # ------------------------------------------------------------------ JSON-LD
    def jsonld(self, p):
        B = self.base
        org_id, site_id = f"{B}/#org", f"{B}/#website"
        countries = [{"@type": "Country", "name": c} for c in self.site["countries"]]
        graph = []
        is_about = p["key"] == "about" and p["loc_key"] == "global"
        if (p["template"] == "home" and p["loc_key"] == "global") or is_about:
            o = self.site["org"]
            graph.append({"@type": "Organization", "@id": org_id, "name": self.site["brand"],
                          "alternateName": self.site["brand_full"], "url": f"{B}/",
                          "logo": {"@type": "ImageObject", "url": f"{B}/assets/brand/afriscan-logo-512.png",
                                   "width": 512, "height": 512},
                          "description": " ".join(o["description"].split()),
                          "disambiguatingDescription": " ".join(o["disambiguating"].split()),
                          "areaServed": countries, "knowsAbout": o["knows_about"], "sameAs": o["same_as"]})
        if p["template"] == "home" and p["loc_key"] == "global":
            graph.append({"@type": "WebSite", "@id": site_id, "url": f"{B}/", "name": self.site["brand"],
                          "alternateName": self.site["brand_full"],
                          "inLanguage": sorted({self.locales[k]["lang"] for k in self.live_locales}),
                          "publisher": {"@id": org_id}})
        wp = {"@type": "WebPage", "@id": f"{p['abs_url']}#webpage", "url": p["abs_url"], "name": p["title"],
              "description": p["description"], "inLanguage": p["lang"], "isPartOf": {"@id": site_id},
              "primaryImageOfPage": {"@type": "ImageObject", "url": p["og_image"]["url"], "width": 1200, "height": 630}}
        if p["template"] == "contact":
            wp["@type"] = ["WebPage", "ContactPage"]
        if p["template"] in ("hub", "law_hub"):
            wp["@type"] = ["WebPage", "CollectionPage"]
        if is_about:
            wp["@type"] = ["WebPage", "AboutPage"]
            wp["mainEntity"] = {"@id": org_id}
        graph.append(wp)
        svc = p["meta"].get("service", {})
        if svc is not False and (p["template"] in ("industry", "solution", "country_home") or svc):
            svc = svc or {}
            node = {"@type": "Service", "@id": f"{p['abs_url']}#service",
                    "name": svc.get("name") or self.label_of(p),
                    "serviceType": svc.get("type") or self.label_of(p),
                    "description": svc.get("description") or p["description"],
                    "provider": {"@id": org_id}, "url": p["abs_url"],
                    "areaServed": ({"@type": "Country", "name": p["loc"]["country_name"]}
                                   if p["loc"]["country"] else countries)}
            graph.append(node)
            wp["about"] = {"@id": node["@id"]}
        if p["crumbs"]:
            graph.append({"@type": "BreadcrumbList", "@id": f"{p['abs_url']}#breadcrumb", "itemListElement": [
                {"@type": "ListItem", "position": i + 1, "name": c["name"], "item": self.base + c["url"]}
                for i, c in enumerate(p["crumbs"])]})
            wp["breadcrumb"] = {"@id": f"{p['abs_url']}#breadcrumb"}
        if p["template"] == "article":
            m = p["meta"]
            graph.append({"@type": "Article", "@id": f"{p['abs_url']}#article", "headline": p["h1"],
                          "description": p["description"], "inLanguage": p["lang"],
                          "datePublished": str(m["published"]), "dateModified": str(m.get("updated", m["published"])),
                          "author": {"@id": org_id}, "publisher": {"@id": org_id},
                          "image": p["og_image"]["url"], "mainEntityOfPage": {"@id": wp["@id"]}})
        if p["template"] == "law_hub":
            wp["lastReviewed"] = str(p["meta"].get("as_of", self.site["law_as_of"]))
        if p["template"] == "law":
            law = self.law_for(p, p["meta"]["law"])
            wp["lastReviewed"] = str(law["last_reviewed"])
            wp["citation"] = [{"@type": "Legislation", "name": i["title"], "legislationIdentifier": i["identifier"],
                               "legislationDate": str(i["date"]), "legislationJurisdiction": law["country_name"],
                               "url": i["url"]} for i in law["instruments"]]
        if p["meta"].get("faq"):
            graph.append({"@type": "FAQPage", "@id": f"{p['abs_url']}#faq", "mainEntity": [
                {"@type": "Question", "name": f["q"],
                 "acceptedAnswer": {"@type": "Answer", "text": re.sub(r"<[^>]+>", "", self.md.renderInline(f["a"]))}}
                for f in p["meta"]["faq"]]})
        doc = json.dumps({"@context": "https://schema.org", "@graph": graph}, ensure_ascii=False)
        return Markup(doc.replace("</", "<\\/"))

    # ------------------------------------------------------------------ assets
    def copy_static(self):
        src = SITE / "static"
        self.asset_urls = {}
        for f in sorted(src.rglob("*")):
            if f.is_dir():
                continue
            rel = f.relative_to(src)
            if rel.parts[:2] in (("assets", "css"), ("assets", "js")):
                data = f.read_bytes()
                name = f"{f.stem}.{sha(data)}{f.suffix}"
                target = self.dist / rel.parent / name
                self.asset_urls[rel.as_posix()] = "/" + (rel.parent / name).as_posix()
            else:
                target = self.dist / rel
                self.asset_urls[rel.as_posix()] = "/" + rel.as_posix()
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(f, target)

    def write_region_js(self):
        """The country-suggestion banner script, only for sections that exist."""
        self.region_js_url = None
        sites = {}
        for lk in self.live_locales:
            loc = self.locales[lk]
            if not loc["country"]:
                continue
            s = sites.setdefault(loc["country"], {"links": []})
            s["links"].append([loc["lang"], self.home_of(lk)["url"], self.i18n[loc["i18n"]]["lang_name"]
                               if loc["country"] == "MZ" else loc["label"]])
        if not sites:
            return
        tpl = (SITE / "templates/region.js.j2").read_text(encoding="utf-8")
        js = tpl.replace("__SITES__", json.dumps(sites, ensure_ascii=False))
        name = f"region.{sha(js)}.js"
        (self.dist / "assets/js").mkdir(parents=True, exist_ok=True)
        (self.dist / "assets/js" / name).write_text(js, encoding="utf-8")
        self.region_js_url = f"/assets/js/{name}"

    # ------------------------------------------------------------------ special pages
    def special_pages(self):
        """404 and thank-you pages per language (not in content/: they carry no copy of their own)."""
        out = []
        variants = [("global", "404", "notfound"), ("global", "thanks", "thanks")]
        if self.home_of("mz-pt"):
            variants += [("mz-pt", "404", "notfound"), ("mz-pt", "obrigado", "thanks")]
        for lk, slug, kind in variants:
            loc = self.locales[lk]
            t = self.i18n[loc["i18n"]]
            url = f"{loc['prefix']}/{slug}"
            p = dict(meta={"cta": False}, body="", src=None, where=f"(generated {url})", loc_key=lk, loc=loc, t=t,
                     slug=slug, url=url, abs_url=self.base + url,
                     out=self.dist / loc["prefix"].lstrip("/") / f"{slug}.html", key=f"{lk}:{slug}",
                     template=kind, status="published", hub=None, section=None, title=t[kind]["title"],
                     description=t[kind]["text"], h1=t[kind]["h1"], lang=loc["lang"], lang2=loc["lang"][:2],
                     noindex=True, alternates=[], crumbs=[], special=True)
            p["region_links"] = [dict(r, current=False) for r in self.region_links(dict(p, key="home"))]
            out.append(p)
        return out

    def render_special(self, p):
        p["og_image"] = self.og_for(p)
        p["jsonld"] = None
        p["cta"] = None
        ctx = self.page_ctx(p)
        ctx["links"] = [q for q in [self.resolve(k, p["loc_key"]) for k in
                                    ("home", "industries", "solutions", "results", "contact")] if q]
        html = self.env.get_template(f"{p['template']}.html.j2").render(**ctx)
        html = re.sub(r"\n\s*\n+", "\n", html)
        p["html"] = html
        p["out"].parent.mkdir(parents=True, exist_ok=True)
        p["out"].write_text(html, encoding="utf-8")

    # ------------------------------------------------------------------ site files
    def git_lastmod(self, *paths):
        try:
            r = subprocess.run(["git", "log", "-1", "--format=%cs", "--", *map(str, paths)], cwd=ROOT,
                               capture_output=True, text=True, check=True)
            dirty = subprocess.run(["git", "status", "--porcelain", "--", *map(str, paths)], cwd=ROOT,
                                   capture_output=True, text=True, check=True).stdout.strip()
            if dirty or not r.stdout.strip():
                return self.today.isoformat()
            return r.stdout.strip()
        except (subprocess.CalledProcessError, FileNotFoundError):
            return self.today.isoformat()

    def write_sitemaps(self):
        by_map = defaultdict(list)
        for p in self.pages:
            if not p["noindex"] and p["meta"].get("sitemap", True):
                by_map[p["loc"]["sitemap"]].append(p)
        imgs = defaultdict(list)
        for page_url, img in self.images.used:
            if img not in imgs[page_url]:
                imgs[page_url].append(img)
        for name, members in sorted(by_map.items()):
            rows = []
            for p in sorted(members, key=lambda p: p["url"]):
                deps = [p["src"]] + ([SITE / f"data/law/{p['meta']['law']}.yaml"] if p["template"] == "law" else [])
                lm = self.git_lastmod(*deps)
                im = "".join(f"\n    <image:image><image:loc>{self.base}{u}</image:loc></image:image>"
                             for u in imgs.get(p["url"], []))
                rows.append(f"  <url>\n    <loc>{p['abs_url']}</loc>\n    <lastmod>{lm}</lastmod>{im}\n  </url>")
            (self.dist / name).write_text(
                '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" '
                'xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">\n' + "\n".join(rows) + "\n</urlset>\n",
                encoding="utf-8")
        (self.dist / "sitemap.xml").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
            + "".join(f"  <sitemap><loc>{self.base}/{n}</loc></sitemap>\n" for n in sorted(by_map))
            + "</sitemapindex>\n", encoding="utf-8")
        self.sitemap_names = sorted(by_map)

    def write_site_files(self):
        (self.dist / "robots.txt").write_text(
            f"User-agent: *\nAllow: /\nDisallow: /geo\n\nSitemap: {self.base}/sitemap.xml\n", encoding="utf-8")
        lines, self.redirect_sources = [], set()
        for src, dst, code in self.redirects:
            target = dst.split("#")[0]
            if not dst.startswith("http") and target not in self.by_url:
                continue
            if src in self.by_url:
                self.err(f"redirect source {src} is also a built page")
            lines.append(f"{src} {dst} {code}")
            self.redirect_sources.add(src)
        (self.dist / "_redirects").write_text("\n".join(lines) + "\n", encoding="utf-8")
        headers = (SITE / "templates/_headers.j2").read_text(encoding="utf-8")
        (self.dist / "_headers").write_text(headers, encoding="utf-8")
        brand.write_icons(self.dist, self.site["theme_color"])

    # ------------------------------------------------------------------ checks
    def visible_text(self, html, main_only=False):
        if main_only:
            m = re.search(r'<main\b[^>]*>(.*)</main>', html, re.S)
            html = m.group(1) if m else html
        html = re.sub(r"(?s)<(script|style|svg)\b.*?</\1>", " ", html)
        attrs = " ¶ ".join(re.findall(r'\b(?:alt|title|aria-label|placeholder)="([^"]*)"', html))
        # block boundaries become a pilcrow, so guard negation never reaches across blocks
        html = re.sub(r"(?i)</(?:p|li|h[1-6]|td|th|dt|dd|summary|figcaption|div|section|header|a|button|option|label|legend)>|<br\s*/?>",
                      " ¶ ", html)
        text = re.sub(r"<[^>]+>", " ", html)
        return htmllib.unescape(" ".join((text + " ¶ " + attrs).split()))

    def guard_text(self, p):
        h = p["html"]
        metas = " ".join(re.findall(r'<meta (?:name|property)="(?:description|og:[a-z:]+|twitter:[a-z:]+)" content="([^"]*)"', h))
        title = " ".join(re.findall(r"<title>(.*?)</title>", h, re.S))
        strings = []

        def walk(v):
            if isinstance(v, str):
                if not v.startswith("http"):
                    strings.append(v)
            elif isinstance(v, dict):
                for k, x in v.items():
                    if k not in ("@id", "@type", "@context"):
                        walk(x)
            elif isinstance(v, list):
                for x in v:
                    walk(x)

        for ld in re.findall(r'<script type="application/ld\+json">(.*?)</script>', h, re.S):
            try:
                walk(json.loads(ld.replace("<\\/", "</")))
            except json.JSONDecodeError:
                pass
        return " ".join([self.visible_text(h), htmllib.unescape(title), htmllib.unescape(metas), " | ".join(strings)])

    def negated(self, text, start):
        window = text[max(0, start - self.rules["negation_window"]):start].lower()
        window = re.split(r"[.!?;:¶]\s|¶", window)[-1]          # only the current sentence counts
        return any(re.search(rf"(?<![\w']){re.escape(n)}(?![\w'])", window) for n in self.rules["negations"])

    def run_guards(self, p, text):
        for chk in self.rules["checks"]:
            for pat in chk["patterns"]:
                for m in re.finditer(pat, text):
                    if chk.get("negatable") and self.negated(text, m.start()):
                        continue
                    snippet = text[max(0, m.start() - 40):m.end() + 40]
                    msg = f"{p['url']}: [{chk['id']}] {chk['message']}: {m.group(0)!r} in “…{snippet}…”"
                    (self.err if chk["level"] == "error" else self.warn)(msg)
                    break
        if p["lang"] == "pt-MZ":
            for level, table in (("error", self.glossary["banned"]), ("warn", self.glossary["ao90"]),
                                 ("warn", self.glossary["warn"])):
                for pat, use in table.items():
                    if m := re.search(pat, text):
                        (self.err if level == "error" else self.warn)(
                            f"{p['url']}: [pt-MZ] {m.group(0)!r}: use {use}")

    def check_structure(self, p):
        h, url, lim = p["html"], p["url"], self.rules["limits"]
        n_h1 = len(re.findall(r"<h1\b", h))
        if n_h1 != 1:
            self.err(f"{url}: needs exactly one <h1> (found {n_h1})")
        if not re.search(r"<title>[^<]{5,}</title>", h):
            self.err(f"{url}: missing <title>")
        if not re.search(r'<meta name="description" content="[^"]{20,}"', h):
            self.err(f"{url}: missing meta description")
        if not re.search(rf'<link rel="canonical" href="{re.escape(p["abs_url"])}">', h):
            self.err(f"{url}: missing or wrong canonical")
        if not re.search(r'<html lang="[a-zA-Z-]+"', h):
            self.err(f"{url}: missing <html lang>")
        for tag in ("og:title", "og:description", "og:image", "og:url"):
            if f'property="{tag}"' not in h:
                self.err(f"{url}: missing {tag}")
        if not p.get("special"):
            tl, dl = len(p["title"]), len(p["description"])
            if tl > lim["title_max"]:
                self.err(f"{url}: title is {tl} characters (max {lim['title_max']}): {p['title']!r}")
            elif tl > lim["title_warn"]:
                self.warn(f"{url}: title is {tl} characters (aim for ≤{lim['title_warn']})")
            if dl < lim["description_min"] or dl > lim["description_max"]:
                self.err(f"{url}: meta description is {dl} characters (keep {lim['description_min']}–{lim['description_max']})")
            elif dl > lim["description_warn"]:
                self.warn(f"{url}: meta description is {dl} characters (aim for ≤{lim['description_warn']})")
        for img in re.findall(r"<img\b[^>]*>", h):
            for a in ("alt=", "width=", "height="):
                if a not in img:
                    self.err(f"{url}: <img> without {a[:-1]}: {img[:120]}")
        for ld in re.findall(r'<script type="application/ld\+json">(.*?)</script>', h, re.S):
            try:
                doc = json.loads(ld.replace("<\\/", "</"))
            except json.JSONDecodeError as e:
                self.err(f"{url}: JSON-LD does not parse: {e}")
                continue
            blob = json.dumps(doc).lower()
            if '"offer' in blob or '"price' in blob or "pricespecification" in blob:
                self.err(f"{url}: JSON-LD contains an Offer or price")
        if "<main" not in h:
            self.err(f"{url}: missing <main>")
        heads = [int(x) for x in re.findall(r"<h([1-6])\b", self.visible_main(h))]
        for a, b in zip(heads, heads[1:]):
            if b > a + 1:
                self.warn(f"{url}: heading jumps from h{a} to h{b}")
                break

    def visible_main(self, h):
        m = re.search(r"<main\b[^>]*>(.*)</main>", h, re.S)
        return m.group(1) if m else h

    def check_links(self, built):
        ids = {}
        for p in built:
            ids[p["url"]] = set(re.findall(r'\bid="([^"]+)"', p["html"]))
        for p in built:
            for attr, val in re.findall(r'\b(href|src)="([^"]*)"', p["html"]):
                self.check_one_link(p, val, ids)
            for srcset in re.findall(r'\bsrcset="([^"]*)"', p["html"]):
                for part in srcset.split(","):
                    self.check_one_link(p, part.strip().split(" ")[0], ids)

    def check_one_link(self, p, val, ids):
        if not val or val.startswith(("http://", "https://", "mailto:", "tel:", "data:")):
            if val.startswith("http://"):
                self.warn(f"{p['url']}: insecure link {val}")
            return
        if val.startswith("key:"):
            self.err(f"{p['url']}: unresolved link {val}")
            return
        if val.startswith("#"):
            if val[1:] and val[1:] not in ids[p["url"]]:
                self.err(f"{p['url']}: link to missing anchor {val}")
            return
        if not val.startswith("/") or val.startswith("//"):
            self.err(f"{p['url']}: links must be root-relative (/…) or absolute: {val}")
            return
        path, _, frag = val.partition("#")
        path = path.split("?")[0]
        if path in self.redirect_sources:
            self.warn(f"{p['url']}: link {val} goes through a redirect; link the target instead")
            return
        target = self.by_url.get(path) or self.special_by_url.get(path)
        if target:
            if frag and frag not in ids.get(target["url"], set()):
                self.err(f"{p['url']}: link {val}: no id '{frag}' on {target['url']}")
            return
        f = self.dist / path.lstrip("/")
        if path.endswith(".html"):
            self.err(f"{p['url']}: link to {val}; use the clean URL without .html")
        elif not (f.exists() and f.is_file()):
            self.err(f"{p['url']}: broken internal link {val}")

    def check_duplicates(self, built):
        seen_t, seen_d = {}, {}
        for p in built:
            if p["noindex"]:
                continue
            for field, seen in (("title", seen_t), ("description", seen_d)):
                v = p[field]
                if v in seen:
                    self.err(f"duplicate {field} on {p['url']} and {seen[v]}: {v!r}")
                seen[v] = p["url"]

    def shingles(self, text, n=5):
        w = re.findall(r"\w+", text.lower())
        return {" ".join(w[i:i + n]) for i in range(max(0, len(w) - n + 1))}

    def check_similarity(self, built):
        lim = self.rules["limits"]
        idx = [p for p in built if not p["noindex"] and not p.get("special")]
        sh = {p["url"]: self.shingles(self.visible_text(p["html"], main_only=True).replace("¶", " ")) for p in idx}
        for i, a in enumerate(idx):
            for b in idx[i + 1:]:
                if a["loc_key"] == b["loc_key"] or a["lang2"] != b["lang2"]:
                    continue
                related = a["key"] == b["key"] or (a["template"] == b["template"] and a["loc"]["country"]
                                                   and b["loc"]["country"] and a["template"] != "contact")
                if not related:
                    continue
                sa, sb = sh[a["url"]], sh[b["url"]]
                if not sa or not sb:
                    continue
                j = len(sa & sb) / len(sa | sb)
                if j >= lim["similarity_error"]:
                    self.err(f"near-duplicate pages {a['url']} ~ {b['url']}: {j:.2f} of 5-word runs shared "
                             f"(limit {lim['similarity_error']:.2f}); write country-specific copy")
                elif j >= lim["similarity_warn"]:
                    self.warn(f"similar pages {a['url']} ~ {b['url']}: {j:.2f} overlap")

    def check_law(self):
        for cc, law in self.law.items():
            where = f"data/law/{cc}.yaml"
            for k in ("country_name", "last_reviewed", "instruments", "regulator"):
                if k not in law:
                    self.err(f"{where}: missing '{k}'")
            for i in law.get("instruments", []):
                for k in ("id", "title", "identifier", "date", "url", "last_checked"):
                    if not i.get(k):
                        self.err(f"{where}: instrument {i.get('id', '?')} missing '{k}'")
                if i.get("url") and not str(i["url"]).startswith("https://"):
                    self.err(f"{where}: instrument {i.get('id')} url must be https")
            lr = law.get("last_reviewed")
            if isinstance(lr, dt.date) and (self.today - lr).days > self.rules["limits"]["law_review_max_days"]:
                self.warn(f"{where}: last reviewed {lr}, over {self.rules['limits']['law_review_max_days']} days ago")
        for p in self.pages:
            if p["template"] == "law":
                if p["t"]["law"]["not_advice"] not in htmllib.unescape(p["html"]):
                    self.err(f"{p['url']}: law page must show the not-legal-advice line")

    # ------------------------------------------------------------------ main
    def check_i18n(self):
        def keys(d, pre=""):
            out = set()
            for k, v in d.items():
                out.add(pre + k)
                if isinstance(v, dict):
                    out |= keys(v, pre + k + ".")
            return out

        sets = {n: keys(d) for n, d in self.i18n.items()}
        allk = set().union(*sets.values())
        for n, s in sets.items():
            for k in sorted(allk - s):
                self.err(f"data/i18n/{n}.yaml: missing key '{k}'")

    def run(self):
        self.check_i18n()
        pages = self.read_pages()
        if self.dist.exists():
            shutil.rmtree(self.dist)
        self.dist.mkdir(parents=True)
        self.font = SITE / "fonts/inter-latin-wght-normal.woff2"
        self.images = Images(SITE / "images", CACHE_DIR / "img", self.dist / "assets/img")
        self.index(pages)
        self.env = Environment(loader=FileSystemLoader(SITE / "templates"), undefined=StrictUndefined,
                               autoescape=True, trim_blocks=True, lstrip_blocks=True)
        self.copy_static()
        self.write_region_js()
        self.navs = {lk: self.nav_for(lk) for lk in self.locales}
        self.footers = {lk: self.footer_for(lk) for lk in self.locales}
        for p in pages:
            try:
                self.render_page(p)
            except ContentError as e:
                self.err(str(e))
                p["html"] = ""
        specials = self.special_pages()
        self.special_by_url = {s["url"]: s for s in specials}
        for s in specials:
            self.render_special(s)
        self.write_sitemaps()
        self.write_site_files()
        built = [p for p in pages if p.get("html")] + specials
        for p in built:
            if not p["noindex"] or p.get("special"):
                self.run_guards(p, self.guard_text(p))
            self.check_structure(p)
        self.check_links(built)
        self.check_duplicates(built)
        self.check_similarity(built)
        self.check_law()
        return self.report(pages, specials)

    def report(self, pages, specials):
        if not self.quiet:
            for w in self.warnings:
                print("WARN ", w)
            for e in self.errors:
                print("ERROR", e)
            by = defaultdict(int)
            for p in pages:
                by[p["loc_key"]] += 1
            secs = ", ".join(f"{k} {v}" for k, v in by.items())
            print(f"{len(pages)} pages ({secs}) + {len(specials)} generated; "
                  f"{len(self.errors)} errors, {len(self.warnings)} warnings -> {self.dist.relative_to(ROOT) if self.dist.is_relative_to(ROOT) else self.dist}")
        return 1 if self.errors else 0


CACHE_DIR = SITE / ".cache"


def selftest():
    """Seed one mistake per guard into a scratch copy of the content and check the build catches each.

    The real content must build clean first; then every seeded mistake must fail the build with the
    expected message."""
    append = lambda sentence: lambda raw: raw.rstrip() + f"\n\n{sentence}\n"
    drop = lambda field: lambda raw: "\n".join(l for l in raw.split("\n") if not l.startswith(field + ":"))
    cases = {
        "price":          ("global", append("Our surveys start at US$ 1,200 per km."), "[pricing]"),
        "currency-word":  ("global", append("A baseline costs 90 000 meticais."), "[pricing]"),
        "rand":           ("global", append("Budget R1 200 000 for the line."), "[pricing]"),
        "pricing-word":   ("global", append("See our pricing for details."), "[pricing]"),
        "placeholder":    ("global", append("Registered office: [OWNER: address]."), "[placeholder]"),
        "tbd":            ("global", append("Response time TBD."), "[placeholder]"),
        "lorem":          ("global", append("Lorem ipsum dolor sit amet."), "[placeholder]"),
        "licensed":       ("global", append("Afridrone is a licensed drone operator in Mozambique."), "[certificates]"),
        "certified":      ("global", append("Our certified pilots fly every survey."), "[certificates]"),
        "we-hold":        ("global", append("We hold the SACAA ROC for this work."), "[certificates]"),
        "roc-holder":     ("global", append("As an ROC holder in Nigeria we can fly anywhere."), "[certificates]"),
        "realtime":       ("global", append("Real-time encroachment alerts for your pipeline."), "[overstatement]"),
        "24-7":           ("global", append("We watch your corridor 24/7."), "[overstatement]"),
        "live":           ("global", append("Live detection of new buildings."), "[overstatement]"),
        "instant":        ("global", append("Instant reports for every route."), "[overstatement]"),
        "cloud":          ("global", append("Our optical imagery sees through cloud."), "[overstatement]"),
        "court":          ("global", append("Our evidence packs are court-grade."), "[overstatement]"),
        "people":         ("global", append("We track people entering the right of way."), "[overstatement]"),
        "accuracy":       ("global", append("Detection accuracy is 97% on African roofs."), "[overstatement]"),
        "coming-soon":    ("global", append("Radar screening is coming soon."), "[overstatement]"),
        "portal":         ("global", append("Log in to the client portal to see results."), "[overstatement]"),
        "br-monitoring":  ("mz-pt", append("Fazemos o monitoramento do gasoduto."), "'monitoramento'"),
        "br-equipe":      ("mz-pt", append("A nossa equipe responde."), "'equipe'"),
        "br-contato":     ("mz-pt", append("Entre em contato conosco."), "'contato'"),
        "broken-link":    ("global", append("See [our method](/no-such-page)."), "broken internal link"),
        "bad-anchor":     ("global", append("See [the FAQ](/faq#no-such-anchor)."), "no id 'no-such-anchor'"),
        "html-link":      ("global", append("See [results](/results.html)."), "without .html"),
        "no-h1":          ("global", drop("h1"), "needs 'h1'"),
        "no-description": ("global", drop("description"), "needs 'description'"),
        "no-title":       ("global", drop("title"), "needs 'title'"),
    }
    (CACHE_DIR / "selftest").mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="run-", dir=CACHE_DIR / "selftest"))
    failed = []
    try:
        clean = Build(dist=tmp / "clean" / "dist", quiet=True)
        code = clean.run()
        print(f"{'PASS' if code == 0 else 'FAIL'}  {'clean build':15s} {len(clean.errors)} errors on the real content")
        if code:
            failed.append("clean build")
            for e in clean.errors[:10]:
                print("      ", e)
        for name, (lk, mutate, expect) in cases.items():
            content = tmp / name / "content"
            shutil.copytree(SITE / "content", content)
            if lk == "mz-pt":
                shutil.copytree(SITE / "tests/fixtures/mz-pt", content / "mz/pt", dirs_exist_ok=True)
                target = content / "mz/pt/index.md"
            else:
                target = content / "global/faq.md"
            target.write_text(mutate(target.read_text(encoding="utf-8")), encoding="utf-8")
            b = Build(content_dir=content, dist=tmp / name / "dist", quiet=True)
            code = b.run()
            hit = [e for e in b.errors if expect in e]
            ok = code == 1 and hit
            print(f"{'PASS' if ok else 'FAIL'}  {name:15s} {(hit or b.errors or ['no error raised'])[0][:105]}")
            if not ok:
                failed.append(name)
        # near-duplicate: an English country page that copies the global one
        content = tmp / "dup" / "content"
        shutil.copytree(SITE / "content", content)
        shutil.copytree(SITE / "tests/fixtures/za-dup", content / "za", dirs_exist_ok=True)
        _, body = split_front_matter((content / "global/results.md").read_text(encoding="utf-8"), "results")
        (content / "za/results.md").write_text(
            "---\nkey: results\ntitle: Sample outputs South Africa | AfriScan\n"
            "description: A copy of the global sample page with a different title, which the build must refuse as a near duplicate.\n"
            "h1: Sample outputs in South Africa\n---\n" + body, encoding="utf-8")
        b = Build(content_dir=content, dist=tmp / "dup" / "dist", quiet=True)
        b.run()
        hit = [e for e in b.errors if "near-duplicate" in e]
        print(f"{'PASS' if hit else 'FAIL'}  {'near-duplicate':15s} {hit[0][:105] if hit else 'no error raised'}")
        if not hit:
            failed.append("near-duplicate")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("selftest:", "all guards fired" if not failed else f"FAILED {failed}")
    return 1 if failed else 0


def demo(out):
    """Build the real content plus the template fixtures in tests/fixtures/demo into OUT (never deploy).
    Fixtures only fill gaps: a real page or law file with the same path always wins, so the demo keeps
    building once a country section has its own pages."""
    def overlay(src, dst):
        for f in sorted(src.rglob("*")):
            t = dst / f.relative_to(src)
            if f.is_dir():
                t.mkdir(parents=True, exist_ok=True)
            elif not t.exists():
                shutil.copy2(f, t)

    work, law = CACHE_DIR / "demo-content", CACHE_DIR / "demo-law"
    for d in (work, law):
        if d.exists():
            shutil.rmtree(d)
    shutil.copytree(SITE / "content", work)
    overlay(SITE / "tests/fixtures/demo/content", work)
    shutil.copytree(SITE / "data/law", law)
    overlay(SITE / "tests/fixtures/demo/law", law)
    return Build(content_dir=work, dist=Path(out), law_dir=law).run()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--drafts", action="store_true", help="also build status: draft pages (never deploy this)")
    ap.add_argument("--selftest", action="store_true", help="check that every guard fails the build")
    ap.add_argument("--dist", default=str(ROOT / "dist"))
    ap.add_argument("--demo", metavar="OUT", help="build content + template fixtures into OUT, for checking templates")
    args = ap.parse_args()
    if args.selftest:
        return selftest()
    if args.demo:
        return demo(args.demo)
    return Build(dist=Path(args.dist), drafts=args.drafts).run()


if __name__ == "__main__":
    sys.exit(main())
