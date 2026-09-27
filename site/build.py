#!/usr/bin/env python3
"""Build afri-scan.com (the global site and every live country section in data/locales.yaml) from site/ into dist/.

    /opt/favhousecheck/.venv/bin/python3 site/build.py            # build + guards, exit 1 on any error
    /opt/favhousecheck/.venv/bin/python3 site/build.py --drafts --dist /tmp/x   # also build draft pages and
                                                                  # draft sections (local preview only; never into a dist/)
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
import unicodedata
from collections import defaultdict
from pathlib import Path

from html.parser import HTMLParser

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
REGIONS = ["Southern", "East", "West", "Central", "North"]        # data/locales.yaml and data/countries.yaml `region`
LOCALE_KEYS = ("content", "prefix", "lang", "hreflang", "og_locale", "i18n", "country", "country_name", "label",
               "short", "sitemap", "status", "region")
# A non-English i18n file may leave these out: English is used, with a WARN (so a new language file
# does not fail the build each time a section or region is added).
I18N_FALLBACK = ("regions", "sites", "region.stay", "region.banner")
LANG_NAMES = {"en": "English", "pt": "Português", "fr": "Français"}   # only for a draft section with no i18n file
THANKS_SLUGS = {"pt": "obrigado", "fr": "merci"}                     # thank-you page per language (English: thanks)
DRAFT_BANNER = "Draft: not published"
TABLE_LABEL = {"pt": "Tabela", "fr": "Tableau"}                       # aria-label of the scrolling table box

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
    "countries": {"allowed": {"cols"}},
}


def load_yaml(p):
    return yaml.safe_load(Path(p).read_text(encoding="utf-8"))


NAME_TOKEN = re.compile(r"[^\W_]+")
NAME_GAP = re.compile(r"[\s.\-\u2010-\u2015/]{1,3}")
_NAME_DIGESTS = {}


def name_digest(name):
    """sha256 of a name as data/rules.yaml's withdrawn_names stores it: lower case, separators
    removed, leading zeros dropped from numbers."""
    joined = "".join(NAME_TOKEN.findall(name.lower()))
    joined = re.sub(r"(?<![0-9])0+(?=[0-9])", "", joined)
    if joined not in _NAME_DIGESTS:
        _NAME_DIGESTS[joined] = hashlib.sha256(joined.encode()).hexdigest()
    return _NAME_DIGESTS[joined]


def name_candidates(text, words=3):
    """Every run of up to WORDS tokens joined only by spaces, full stops, hyphens, dashes or slashes."""
    text = text.lower()
    toks = [(m.start(), m.end()) for m in NAME_TOKEN.finditer(text)]
    for i, (a, b) in enumerate(toks):
        end = b
        for j in range(i, min(i + words, len(toks))):
            if j > i:
                if not NAME_GAP.fullmatch(text[toks[j - 1][1]:toks[j][0]]):
                    break
                end = toks[j][1]
            yield a, end, text[a:end]


def sha(data, n=8):
    return hashlib.sha256(data if isinstance(data, bytes) else data.encode()).hexdigest()[:n]


def checkout_dists():
    """dist/ of this checkout and of every git worktree of the repository."""
    roots = {ROOT}
    try:
        out = subprocess.run(["git", "-c", "safe.directory=*", "worktree", "list", "--porcelain"], cwd=ROOT,
                             capture_output=True, text=True, check=True, timeout=30).stdout
        roots |= {Path(line.split(" ", 1)[1]) for line in out.splitlines() if line.startswith("worktree ")}
    except (subprocess.SubprocessError, FileNotFoundError, OSError):
        pass
    return {(r / "dist").resolve() for r in roots}


def dist_refusal(dist):
    """Why a drafts build must not write to DIST (a checkout's dist/, or anything inside one), else None."""
    d = Path(dist).resolve()
    for cd in sorted(checkout_dists()):
        if d == cd or cd in d.parents:
            return f"{d} is inside {cd}, a checkout's dist/"
    for anc in (d, *d.parents):
        if anc.name == "dist" and (anc.parent / "site" / "build.py").exists():
            return f"{d} is inside {anc}, the dist/ of a checkout at {anc.parent}"
    return None


class PageRefs(HTMLParser):
    """Every URL-valued attribute of a built page, with what holds it (the country menu, a country-sites
    button, a /countries entry, the header or footer, an hreflang alternate), plus the JSON-LD blocks."""
    HOLDERS = (("region", "region-menu item"), ("country-sites", "country-sites button"),
               ("countries-directory", "/countries entry"))
    VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack, self.refs, self.jsonld, self._ld = [], [], [], None

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        classes = (a.get("class") or "").split()
        holder = next((what for cls, what in self.HOLDERS if cls in classes), None)
        if not holder and tag in ("header", "footer"):
            holder = f"{tag} link"
        where = holder or next((h for _, h in reversed(self.stack) if h), "internal link")
        if tag not in self.VOID:
            self.stack.append((tag, holder))
        if tag == "script" and a.get("type") == "application/ld+json":
            self._ld = []
        if tag == "link" and (a.get("rel") or "").lower() == "alternate" and a.get("hreflang"):
            self.refs.append(("hreflang alternate", a.get("href") or "", a["hreflang"]))
            return
        for k in ("href", "src", "action", "poster", "content", "data-href"):
            if a.get(k):
                self.refs.append(("meta tag" if tag == "meta" else where, a[k], None))
        for k in ("srcset", "imagesrcset"):
            for part in (a.get(k) or "").split(","):
                if part.strip():
                    self.refs.append((where, part.strip().split(" ")[0], None))

    def handle_endtag(self, tag):
        if tag == "script" and self._ld is not None:
            self.jsonld.append("".join(self._ld))
            self._ld = None
        for i in range(len(self.stack) - 1, -1, -1):
            if self.stack[i][0] == tag:
                del self.stack[i:]
                break

    def handle_data(self, data):
        if self._ld is not None:
            self._ld.append(data)


class Build:
    def __init__(self, content_dir=SITE / "content", dist=ROOT / "dist", drafts=False, quiet=False,
                 today=None, law_dir=SITE / "data/law"):
        self.content_dir, self.dist, self.drafts, self.quiet = content_dir, dist, drafts, quiet
        self.today = today or dt.date.today()
        self.errors, self.warnings = [], []
        self.site = load_yaml(SITE / "data/site.yaml")
        self.base = self.site["base_url"].rstrip("/")
        self.locales = load_yaml(SITE / "data/locales.yaml")
        self.countries = load_yaml(SITE / "data/countries.yaml")
        # a draft section may not have its i18n file yet (it then falls back to en.yaml; check_locales)
        self.i18n = {n: load_yaml(SITE / f"data/i18n/{n}.yaml")
                     for n in dict.fromkeys(l.get("i18n", "en") for l in self.locales.values())
                     if (SITE / f"data/i18n/{n}.yaml").exists()}
        self.catalogue = load_yaml(SITE / "data/catalogue.yaml")
        self.rules = load_yaml(SITE / "data/rules.yaml")
        self.glossary = load_yaml(SITE / "data/glossary/pt-MZ.yaml")
        self.redirects = load_yaml(SITE / "data/redirects.yaml")
        self.reviews = load_yaml(SITE / "data/reviews.yaml")
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

    # ------------------------------------------------------------------ sections and countries
    def is_live(self, lk):
        return self.locales[lk].get("status") == "live"

    def section_built(self, lk):
        """Live sections always; draft sections only in a --drafts build."""
        return self.is_live(lk) or self.drafts

    def t_of(self, lk):
        """UI strings of section lk (en.yaml for a draft section whose i18n file does not exist yet)."""
        return self.i18n.get(self.locales[lk]["i18n"]) or self.i18n["en"]

    def tr(self, lk, *path):
        """A UI string for section lk, falling back to English for the I18N_FALLBACK keys."""
        for t in (self.t_of(lk), self.i18n["en"]):
            v = t
            for k in path:
                v = v.get(k) if isinstance(v, dict) else None
            if v is not None:
                return v
        return None

    def cat_lang(self, lang2):
        """The data/catalogue.yaml language key for a page language: its own when every industry, solution
        and service group has it (en, pt; fr once written), else en."""
        if not hasattr(self, "_cat_langs"):
            dicts = [d for c in self.catalogue["industries"] + self.catalogue["solutions"]
                     for d in (c["name"], c["blurb"], c.get("menu") or c["blurb"])]
            dicts += [g["name"] for g in self.catalogue.get("service_groups", []) + self.catalogue.get("solution_groups", [])
                      if isinstance(g.get("name"), dict)]
            self._cat_langs = set.intersection(*(set(d) for d in dicts)) if dicts else {"en"}
        return lang2 if lang2 in self._cat_langs else "en"

    def lang_name(self, lk):
        loc = self.locales[lk]
        if loc["i18n"] in self.i18n:
            return self.i18n[loc["i18n"]]["lang_name"]
        return LANG_NAMES.get(loc["lang"][:2], loc["lang"])

    def country_of(self, lk):
        cc = self.locales[lk].get("country")
        return self.countries.get(cc.lower()) if cc else None

    def section_of_path(self, path):
        """The section an URL path belongs to: the longest prefix that matches at a / boundary."""
        best, n = "global", 0
        for lk, loc in self.locales.items():
            pre = loc["prefix"]
            if pre and (path == pre or path.startswith(pre + "/")) and len(pre) > n:
                best, n = lk, len(pre)
        return best

    def reserved_segments(self, lk):
        """First slug segments that belong to a nested section: "pt" under /mz/, "fr" under /cd/, every
        country prefix under the global site. Draft sections count, so no page can take their URLs."""
        pre, out = self.locales[lk]["prefix"], {}
        for k, loc in self.locales.items():
            if k != lk and loc["prefix"].startswith(pre + "/"):
                out.setdefault(loc["prefix"][len(pre) + 1:].split("/")[0], k)
        return out

    def check_locales(self):
        """data/locales.yaml against data/countries.yaml and data/i18n/."""
        seen_prefix, seen_code = {}, {}
        for lk, loc in self.locales.items():
            where = f"data/locales.yaml: {lk}"
            missing = [k for k in LOCALE_KEYS if k not in loc]
            if missing:
                self.err(f"{where}: missing {', '.join(missing)}")
                continue
            if loc["status"] not in ("live", "draft"):
                self.err(f"{where}: status must be live or draft")
            if lk == "global":
                if loc["status"] != "live" or loc["region"] is not None or loc["country"]:
                    self.err(f"{where}: the global section is live, with region: null and country: null")
            elif loc["region"] not in REGIONS:
                self.err(f"{where}: region must be one of {REGIONS}")
            if loc["prefix"] in seen_prefix:
                self.err(f"{where}: prefix {loc['prefix']!r} is also {seen_prefix[loc['prefix']]}'s")
            seen_prefix[loc["prefix"]] = lk
            if loc["prefix"] and not re.fullmatch(r"(?:/[a-z]{2})+", loc["prefix"]):
                self.err(f"{where}: prefix must look like /xx or /xx/yy")
            for code in loc["hreflang"]:
                if code in seen_code:
                    self.err(f"{where}: hreflang {code} is also carried by {seen_code[code]} (one section per code)")
                seen_code[code] = lk
            if loc["hreflang"][0] != (loc["lang"] if lk != "global" else "en"):
                self.err(f"{where}: the first hreflang code must be the section's own ({loc['lang']})")
            if loc["i18n"] not in self.i18n:
                if loc["status"] == "live":
                    self.err(f"{where}: a live section needs data/i18n/{loc['i18n']}.yaml")
                elif self.drafts and (self.content_dir / loc["content"]).exists() and any(
                        (self.content_dir / loc["content"]).rglob("*.md")):
                    self.warn(f"{where}: no data/i18n/{loc['i18n']}.yaml yet; the draft uses English UI strings")
            if lk != "global":
                c = self.country_of(lk)
                if not c:
                    self.err(f"{where}: country {loc['country']} is not in data/countries.yaml")
                elif c.get("region") != loc["region"]:
                    self.err(f"{where}: region {loc['region']} but data/countries.yaml has {c.get('region')}")
            if lk not in (self.i18n["en"].get("sites") or {}):
                self.err(f"data/i18n/en.yaml: sites.{lk} is missing (the section's name in menus)")
        for r in REGIONS:
            if not (self.i18n["en"].get("regions") or {}).get(r):
                self.err(f"data/i18n/en.yaml: regions.{r} is missing")
        live_cc = {self.locales[lk]["country"].lower() for lk in self.locales
                   if self.locales[lk].get("country") and self.is_live(lk)}
        for cc, c in (self.countries or {}).items():
            where = f"data/countries.yaml: {cc}"
            if not re.fullmatch(r"[a-z]{2}", str(cc)):
                self.err(f"{where}: keys are ISO 3166-1 alpha-2 codes in lower case")
            if not c.get("name") or c.get("region") not in REGIONS:
                self.err(f"{where}: needs a name and a region in {REGIONS}")
            if c.get("status") not in ("live", "launching", "research"):
                self.err(f"{where}: status must be live, launching or research")
            langs = c.get("languages")
            if langs is not None and not (isinstance(langs, list) and all(re.fullmatch(r"[a-z]{2}", str(x)) for x in langs)):
                self.err(f"{where}: languages is a list of ISO 639-1 codes, or null")
            tzs = c.get("timezones") or []
            if not all(re.fullmatch(r"[A-Z][A-Za-z_]+/[A-Za-z_\-]+", str(z)) for z in tzs):
                self.err(f"{where}: timezones are IANA names (Africa/Lusaka)")
            if c.get("status") == "live" and cc not in live_cc:
                self.err(f"{where}: status live, but no section in data/locales.yaml for {cc.upper()} is live")
            elif cc in live_cc and c.get("status") != "live":
                self.warn(f"{where}: a {cc.upper()} section is live; set status: live")

    # ------------------------------------------------------------------ pages
    def read_pages(self):
        pages = []
        for lk, loc in self.locales.items():
            if not self.section_built(lk):
                continue
            base = self.content_dir / loc["content"]
            if not base.exists():
                continue
            nested = [Path(l["content"]) for k, l in self.locales.items()
                      if k != lk and Path(l["content"]).is_relative_to(Path(loc["content"]))]
            for f in sorted(base.rglob("*.md")):
                rel = f.relative_to(base)
                if any(f.relative_to(self.content_dir).is_relative_to(n) for n in nested):
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
        nested = self.reserved_segments(lk).get(slug.split("/")[0])
        if nested:
            raise ContentError(f"{where}: slug {slug!r} is reserved: {self.locales[nested]['prefix']}/ is the "
                               f"{nested} section")
        if slug.split("/")[0] in {"assets", "geo", "404", "thanks", *THANKS_SLUGS.values()}:
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
        draft = status != "published" or not self.is_live(lk)
        return dict(meta=meta, body=body, src=f, where=where, loc_key=lk, loc=loc, t=self.t_of(lk),
                    slug=slug, url=url, abs_url=self.base + url, out=out, key=key, template=template,
                    status=status, hub=hub, section=section, title=meta["title"].strip(),
                    description=" ".join(str(meta["description"]).split()), h1=meta["h1"].strip(),
                    lang=loc["lang"], lang2=loc["lang"][:2], draft=draft, noindex=bool(meta.get("noindex")) or draft)

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
        lang = self.cat_lang(p["lang2"])
        cat = self.cat_ind.get(p["key"]) or self.cat_sol.get(p["key"])
        if cat:
            return cat["name"][lang]
        return p["h1"]

    def label_in(self, q, lk, text=False):
        """How page q is named where section lk links to it. A page in another language takes its label
        (and card text) from that section's i18n `foreign_pages.<key>`; a missing entry fails the build,
        so no English label reaches a Portuguese menu, footer, breadcrumb or card unnoticed."""
        loc = self.locales[lk]
        if q["lang2"] == loc["lang"][:2]:
            return (self.label_of(q), q["meta"].get("summary") or q["description"]) if text else self.label_of(q)
        if loc["i18n"] not in self.i18n:            # a draft section with no i18n file yet: own labels
            return (self.label_of(q), q["meta"].get("summary") or q["description"]) if text else self.label_of(q)
        fp = (self.i18n[loc["i18n"]].get("foreign_pages") or {}).get(q["key"]) or {}
        if not fp.get("label") or not fp.get("blurb"):
            self.err(f"data/i18n/{loc['i18n']}.yaml: foreign_pages.{q['key']} needs a label and a blurb "
                     f"({q['url']} is linked from {lk} pages)")
        label = fp.get("label") or self.label_of(q)
        return (label, fp.get("blurb") or q["description"]) if text else label

    def site_label(self, lk, viewer_lk):
        return self.tr(viewer_lk, "sites", lk) or self.locales[lk]["label"]

    def by_region(self, items, viewer_lk, key="section"):
        """Items that carry a section key, as (items with no region, i.e. global) and region groups in REGIONS
        order. Within a region, countries sort by how the viewer's language names them (the label of the
        country's first section in data/locales.yaml, accents folded); a country's sections stay together,
        in data/locales.yaml order."""
        order = list(self.locales)
        first = {}
        for k in order:
            first.setdefault(self.locales[k]["country"], k)
        fold = lambda s: "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c)).casefold()
        top, groups = [], defaultdict(list)
        for it in items:
            region = self.locales[it[key]].get("region")
            (groups[region] if region else top).append(it)
        out = []
        for r in REGIONS:
            if groups[r]:
                its = sorted(groups[r], key=lambda it: (fold(self.site_label(first[self.locales[it[key]]["country"]], viewer_lk)),
                                                        order.index(it[key])))
                out.append({"key": r.lower(), "region": r, "title": self.tr(viewer_lk, "regions", r), "items": its})
        return top, out

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
        if isinstance(d, int) or (isinstance(d, str) and re.fullmatch(r"\d{4}", d)):
            return str(d)                       # year only, e.g. an Act known by its year and number
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
        # sections with a home in this build (live ones; draft ones too in a --drafts build), and the
        # live ones among them: only those are ever suggested by the banner or named in JSON-LD
        self.live_locales = [lk for lk in self.locales if self.home_of(lk)]
        self.public_locales = [lk for lk in self.live_locales if self.is_live(lk)]
        for p in pages:
            p["alternates"] = self.alternates(p)
            p["region_links"] = self.region_links(p)
            p["region_menu"] = self.region_menu(p)
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
            links.append({"key": lk, "section": lk, "label": self.site_label(lk, p["loc_key"]), "lang": loc["lang"],
                          "href": target["url"], "current": lk == p["loc_key"], "draft": not self.is_live(lk)})
        return links

    def region_menu(self, p):
        top, groups = self.by_region(p["region_links"], p["loc_key"])
        return {"top": top, "groups": groups}

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
                crumbs.append({"name": self.label_in(hub, lk), "url": hub["url"]})
        parent = p["meta"].get("parent")
        if parent:
            pp = self.resolve(parent, lk)
            if not pp:
                self.err(f"{p['url']}: parent '{parent}' is not a built page key")
            elif pp["url"] not in {c["url"] for c in crumbs}:
                crumbs.append({"name": self.label_in(pp, lk), "url": pp["url"]})
        crumbs.append({"name": self.label_of(p), "url": p["url"]})
        return crumbs

    # ------------------------------------------------------------------ navigation
    def nav_for(self, lk):
        page_lang = self.locales[lk]["lang"]
        lang = self.cat_lang(page_lang[:2])
        t = self.t_of(lk)

        def item(p, label=None, blurb=None, **kw):
            foreign = p["lang2"] != page_lang[:2]
            if foreign and label is None:           # a page in another language: label from i18n foreign_pages
                label, fblurb = self.label_in(p, lk, text=True)
                blurb = fblurb if blurb is None else blurb
            d = {"label": label or self.label_of(p), "href": p["url"],
                 "blurb": blurb if blurb is not None else p["meta"].get("nav_blurb", ""),
                 "lang": p["lang"] if foreign else None,
                 "badge": t["nav"]["lang_badge"] if foreign and label and kw.get("key") else None}
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
                     and page_lang[:2] in (p["meta"].get("nav_langs") or [page_lang[:2]])]
            cands.sort(key=lambda p: (p["meta"].get("nav_order", 50), p["title"]))
            for p in cands:
                if p["loc_key"] == lk or (p["loc_key"] == "global" and not self.find(p["key"], lk)):
                    if p["key"] in seen:
                        continue
                    seen.add(p["key"])
                    items.append(item(p))
            extra = {}
            if gname == "countries":
                homes = [item(self.home_of(k), self.site_label(k, lk), "", lang=None if self.locales[k]["lang"][:2] == page_lang[:2] else self.locales[k]["lang"],
                              section=k, draft=not self.is_live(k))
                         for k in self.live_locales if k != "global"]
                extra = {"regions": self.by_region(homes, lk)[1], "pages": items}   # the panel groups homes by region
                items = homes + items
            hub = self.hub_page(gname, lk)
            if items or hub:
                groups.append({"key": gname, "label": t["nav"][gname],
                               "href": (hub or {"url": items[0]["href"]})["url"], "items": items, **extra})
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
        html = wrap_tables(html, TABLE_LABEL.get(p["lang2"], "Table"))
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
        lang = self.cat_lang(p["lang2"])
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
        lang = self.cat_lang(p["lang2"])
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
            name = (s.get(f"name_{p['lang2']}") if p["lang2"] != "en" else None) or s["name"]
            line = (s.get(f"line_{p['lang2']}") if p["lang2"] != "en" else None) or s["line"]
            out.append({"id": s["id"], "name": name, "line": line, "group": s["group"]})
        return out

    def cmp_catalogue(self, p, b, ctx):
        lang = self.cat_lang(p["lang2"])
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
        cards = [{"title": self.label_in(q, lk), "text": self.label_in(q, lk, text=True)[1], "href": q["url"],
                  "icon": q["meta"].get("icon"), "date": q["meta"].get("published")} for q in items]
        return {"cards": cards}

    def law_for(self, p, cc):
        """data/law/<cc>.yaml with each instrument's optional title_<lang> / identifier_<lang> / note_<lang>
        (title_pt on PT pages, title_fr on FR pages…) used instead of the English wording."""
        law, sfx = self.law.get(cc), "_" + p["lang2"]
        if not law or p["lang2"] == "en":
            return law
        loc = lambda i: {**i, **{k: i[k + sfx] for k in ("title", "identifier", "note") if i.get(k + sfx)}}
        return {**law, "country_name": law.get("country_name" + sfx) or law["country_name"],
                "instruments": [loc(i) for i in law.get("instruments", [])]}

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
            sites.append({"section": k, "label": self.site_label(k, p["loc_key"]), "href": target["url"],
                          "lang": self.locales[k]["lang"], "draft": not self.is_live(k)})
        return {"sites": sites, "groups": self.by_region(sites, p["loc_key"])[1]}

    def cmp_countries(self, p, b, ctx):
        """The /countries directory: one card per country with a site, grouped by region."""
        lk = p["loc_key"]
        by_cc = {}
        for k in self.live_locales:
            if self.locales[k]["country"]:
                by_cc.setdefault(self.locales[k]["country"], []).append(k)
        entries = []
        for cc, secs in by_cc.items():
            same = next((k for k in secs if self.locales[k]["lang"][:2] == p["lang2"]), None)
            home = self.home_of(same or secs[0])          # the reader's language where the country has it
            title = (self.country_of(secs[0]) or {}).get("name") if p["lang2"] == "en" else None
            title = title or (self.locales[same]["country_name"] if same else self.site_label(secs[0], lk))
            entries.append({"section": secs[0], "title": title, "href": home["url"], "lang": home["lang"],
                            "text": home["meta"].get("summary") or home["description"],
                            "langs": " · ".join(self.lang_name(k) for k in secs),
                            "links": [{"label": self.lang_name(k), "href": self.home_of(k)["url"],
                                       "lang": self.locales[k]["lang"]} for k in secs] if len(secs) > 1 else [],
                            "draft": not all(self.is_live(k) for k in secs)})
        return {"groups": self.by_region(entries, lk)[1], "cols": b.attrs.get("cols", "4")}

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
                    lang_key=self.cat_lang(p["lang2"]), catalogue=self.catalogue,
                    resolve=lambda k: self.resolve(k, lk), footer=self.footers[lk], today=self.today,
                    region_js=self.region_js_url, thanks_url=self.thanks_url(p),
                    label_of=lambda q: self.label_in(q, lk))

    def own_specials(self, lk):
        """A section gets its own 404 and thank-you pages when it has its own (non-English) UI strings."""
        loc = self.locales[lk]
        return lk != "global" and loc["i18n"] != "en" and loc["i18n"] in self.i18n and bool(self.home_of(lk))

    def thanks_slug(self, lk):
        return THANKS_SLUGS.get(self.locales[lk]["lang"][:2], "thanks")

    def thanks_url(self, p):
        lk = p["loc_key"]
        if self.own_specials(lk):
            return f"{self.base}{self.locales[lk]['prefix']}/{self.thanks_slug(lk)}"
        return f"{self.base}/thanks"

    def footer_for(self, lk):
        nav = self.navs[lk]
        g = {x["key"]: x for x in nav}
        explore = [{"label": x["label"], "href": x["href"], "lang": None, "badge": None} for x in nav]
        sites = [{"section": k, "label": self.site_label(k, lk), "href": self.home_of(k)["url"],
                  "lang": self.locales[k]["lang"], "draft": not self.is_live(k)} for k in self.live_locales]
        top, regions = self.by_region(sites, lk)
        return {"explore": explore,
                "industries": g.get("industries", {}).get("items", []),
                "solutions": self.footer_solutions(g.get("solutions", {}).get("items", [])),
                "how": g.get("how", {}).get("items", []),
                "resources": g.get("resources", {}).get("items", []),
                "countries": sites, "countries_top": top, "country_regions": regions}

    def footer_solutions(self, items):
        """Solutions marked `footer: true` in catalogue.yaml, in catalogue order; else the first eight."""
        flagged = {s["key"] for s in self.catalogue["solutions"] if s.get("footer")}
        return [i for i in items if i.get("key") in flagged] if flagged else items[:8]

    def related_cards(self, p, keys):
        lang = self.cat_lang(p["lang2"])
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
            cards.append({"title": cat["name"][lang] if cat else self.label_in(tp, p["loc_key"]),
                          "text": cat["blurb"][lang] if cat else self.label_in(tp, p["loc_key"], text=True)[1],
                          "href": tp["url"], "icon": (cat or {}).get("icon") or tp["meta"].get("icon")})
        return cards

    def guides_for(self, p, skip_hrefs=()):
        """Articles whose `about` lists this page's key, in the page's language: this section's first, then global
        ones; global pages also list the country guides."""
        if p["template"] in ("article", "home", "hub", "contact") or p["noindex"]:
            return []
        found = []
        for a in self.pages:
            if (a["template"] != "article" or a["noindex"] or a is p or a["lang2"] != p["lang2"]
                    or a["url"] in skip_hrefs or p["key"] not in (a["meta"].get("about") or [])):
                continue
            if a["loc_key"] == p["loc_key"]:
                rank = 0
            elif a["loc_key"] == "global":
                rank = 1
            elif p["loc_key"] == "global":
                rank = 2
            else:
                continue
            found.append((rank, a["meta"].get("nav_order", 50), a["title"], a))
        found.sort(key=lambda x: x[:3])
        return [{"title": self.label_in(a, p["loc_key"]), "text": a["meta"].get("summary") or a["description"], "href": a["url"],
                 "icon": a["meta"].get("icon"), "eyebrow": a["meta"].get("eyebrow")} for *_, a in found[:6]]

    def country_row(self, p):
        """Global solution pages: one button per country site, to that country's version of the service or its home."""
        if p["template"] != "solution" or p["loc_key"] != "global":
            return []
        exclude = (self.cat_sol.get(p["key"]) or {}).get("exclude", [])
        sites = []
        for k in self.live_locales:
            if k == "global" or k in exclude:
                continue
            target = self.find(p["key"], k)
            if not target or target["noindex"]:
                target = self.home_of(k)
            sites.append({"section": k, "label": self.site_label(k, p["loc_key"]), "href": target["url"],
                          "lang": self.locales[k]["lang"], "draft": not self.is_live(k)})
        return self.by_region(sites, p["loc_key"])[1]

    def written_for(self, p):
        """Articles: the industry pages (in the article's own section and language) that its `about` keys name."""
        if p["template"] != "article":
            return []
        lang = self.cat_lang(p["lang2"])
        out = []
        for k in p["meta"].get("about") or []:
            tp = self.resolve(k, p["loc_key"])
            if tp and tp["template"] == "industry" and not tp["noindex"] and tp["lang2"] == p["lang2"]:
                label = self.cat_ind[k]["name"][lang] if k in self.cat_ind else self.label_of(tp)
                out.append({"label": label, "href": tp["url"]})
        return out

    def used_in(self, p):
        lang = self.cat_lang(p["lang2"])
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
        for k in m.get("about") or []:
            if k not in self.by_key:
                self.err(f"{p['url']}: about key '{k}' is not a built page")
        p["guides"] = self.guides_for(p, {c["href"] for c in p["related"]})
        p["country_row"] = self.country_row(p)
        p["written_for"] = self.written_for(p)
        tones = re.findall(r'<section class="section section--(\w+)', str(p["body_html"]))
        last = tones[-1] if tones else ("dark" if p["template"] in ("home", "country_home") else "light")
        flip = lambda t: "light" if t == "alt" else "alt"
        p["related_tone"] = flip(last)
        has_related = p["related"] or p["guides"] or p["country_row"]
        p["faq_tone"] = flip(p["related_tone"] if has_related else last)
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
        try:
            name = brand.card(self.dist / "assets/og", p["url"], head, sub, region, self.font, CACHE_DIR / "og")
        except brand.CardTextError as e:
            self.err(f"{p['url']}: {e} (og.headline / og.subline in the front matter)")
            return {"url": f"{self.base}/assets/og/missing.jpg", "alt": og.get("alt") or head}
        return {"url": f"{self.base}/assets/og/{name}", "alt": og.get("alt") or head}

    # ------------------------------------------------------------------ JSON-LD
    def served_countries(self):
        """English names of the countries with a live section, in data/locales.yaml order."""
        names = []
        for lk in self.public_locales:
            c = self.country_of(lk)
            if c and c["name"] not in names:
                names.append(c["name"])
        return names

    def jsonld(self, p):
        B = self.base
        org_id, site_id = f"{B}/#org", f"{B}/#website"
        countries = [{"@type": "Country", "name": c} for c in self.served_countries()]
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
                          "inLanguage": sorted({self.locales[k]["lang"] for k in self.public_locales}),
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

    def banner_sections(self):
        """Sections the country banner may suggest: live ones with a home. Never a draft, even in --drafts."""
        return [lk for lk in self.public_locales if self.locales[lk]["country"]]

    def write_region_js(self):
        """The country-suggestion banner script, generated from the live country sections."""
        self.region_js_url = None
        by_cc = {}
        for lk in self.banner_sections():
            by_cc.setdefault(self.locales[lk]["country"], []).append(lk)
        sites, tz, ui = {}, {}, {}
        for cc, secs in by_cc.items():
            first = self.locales[secs[0]]
            # the banner speaks the country's first language only when that i18n file has its own sentence
            own = (self.i18n.get(first["i18n"], {}).get("region") or {}).get("banner")
            text = (own or self.i18n["en"]["region"]["banner"]).format(
                country=first["country_name"] if own else (self.country_of(secs[0]) or {}).get("name", first["country_name"]))
            sites[cc] = {"lang": first["lang"] if own else "en", "text": text,
                         "links": [[self.locales[k]["lang"], self.home_of(k)["url"],
                                    self.lang_name(k) if len(secs) > 1 else self.locales[k]["label"]] for k in secs]}
            for z in (self.country_of(secs[0]) or {}).get("timezones") or []:
                tz[z] = cc
        ui["en"] = {"bar": self.i18n["en"]["region"]["label"], "stay": self.i18n["en"]["region"]["stay"]}
        for lk in self.public_locales:
            loc = self.locales[lk]
            ui[loc["lang"]] = {"bar": self.tr(lk, "region", "label"), "stay": self.tr(lk, "region", "stay")}
        if not sites:
            return
        tpl = (SITE / "templates/region.js.j2").read_text(encoding="utf-8")
        data = {"sites": sites, "tz": tz, "ui": ui}
        js = tpl.replace("__DATA__", json.dumps(data, ensure_ascii=False, sort_keys=True))
        name = f"region.{sha(js)}.js"
        (self.dist / "assets/js").mkdir(parents=True, exist_ok=True)
        (self.dist / "assets/js" / name).write_text(js, encoding="utf-8")
        self.region_js_url = f"/assets/js/{name}"

    # ------------------------------------------------------------------ special pages
    def special_pages(self):
        """404 and thank-you pages per language (not in content/: they carry no copy of their own)."""
        out = []
        variants = [("global", "404", "notfound"), ("global", "thanks", "thanks")]
        for lk in self.live_locales:
            if self.own_specials(lk):
                variants += [(lk, "404", "notfound"), (lk, self.thanks_slug(lk), "thanks")]
        for lk, slug, kind in variants:
            loc = self.locales[lk]
            t = self.t_of(lk)
            url = f"{loc['prefix']}/{slug}"
            p = dict(meta={"cta": False}, body="", src=None, where=f"(generated {url})", loc_key=lk, loc=loc, t=t,
                     slug=slug, url=url, abs_url=self.base + url,
                     out=self.dist / loc["prefix"].lstrip("/") / f"{slug}.html", key=f"{lk}:{slug}",
                     template=kind, status="published", hub=None, section=None, title=t[kind]["title"],
                     description=t[kind]["text"], h1=t[kind]["h1"], lang=loc["lang"], lang2=loc["lang"][:2],
                     draft=not self.is_live(lk), noindex=True, alternates=[], crumbs=[], special=True)
            p["region_links"] = [dict(r, current=False) for r in self.region_links(dict(p, key="home"))]
            p["region_menu"] = self.region_menu(p)
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
        files = ["/" + f.relative_to(self.dist).as_posix() for f in self.dist.rglob("*") if f.is_file()]
        for src, dst, code in self.redirects:
            target = dst.split("#")[0]
            if not dst.startswith("http") and target not in self.by_url:
                continue
            if src in self.by_url:
                self.err(f"redirect source {src} is also a built page")
            hidden = [f for f in files if (f.startswith(src[:-1]) if src.endswith("*") else f == src)]
            if hidden:
                self.err(f"redirect source {src} would hide {len(hidden)} published file(s), e.g. {hidden[0]}")
            lines.append(f"{src} {dst} {code}")
            self.redirect_sources.add(src)
        (self.dist / "_redirects").write_text("\n".join(lines) + "\n", encoding="utf-8")
        # functions/_middleware.js (host redirects) runs on every route except these, which saves quota
        (self.dist / "_routes.json").write_text(json.dumps(
            {"version": 1, "include": ["/*"], "exclude": ["/assets/*"]}, indent=2) + "\n", encoding="utf-8")
        headers = (SITE / "templates/_headers.j2").read_text(encoding="utf-8")
        langs = "\n".join(f"{self.locales[lk]['prefix']}/*\n  Content-Language: {self.locales[lk]['lang']}"
                          for lk in self.public_locales if lk != "global" and self.locales[lk]["lang"][:2] != "en")
        (self.dist / "_headers").write_text(headers.replace("__CONTENT_LANGUAGE__", langs), encoding="utf-8")
        brand.write_icons(self.dist, self.site["theme_color"])
        key = str(self.site.get("indexnow_key", ""))
        if key:
            if not re.fullmatch(r"[A-Za-z0-9-]{8,128}", key):
                self.err(f"indexnow_key {key!r}: 8-128 letters, digits or dashes")
            else:
                (self.dist / f"{key}.txt").write_text(key, encoding="utf-8")

    # ------------------------------------------------------------------ checks
    def visible_text(self, html, main_only=False):
        if main_only:
            m = re.search(r'<main\b[^>]*>(.*)</main>', html, re.S)
            html = m.group(1) if m else html
        html = re.sub(r"(?s)<(script|style)\b.*?</\1>", " ", html)     # inline <svg> text is checked too
        attrs = " ¶ ".join(re.findall(r'\b(?:alt|title|aria-label|placeholder)="([^"]*)"', html))
        # block boundaries become a pilcrow, so guard negation never reaches across blocks
        html = re.sub(r"(?i)</(?:p|li|h[1-6]|td|th|dt|dd|summary|figcaption|div|section|header|a|button|option|label|legend"
                      r"|svg|text|tspan|title|desc)>|<br\s*/?>", " ¶ ", html)
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

    def is_law_page(self, p):
        return p.get("template") == "law" or p.get("key") in (self.rules.get("law_keys") or [])

    def in_scope(self, p, scope):
        if scope in (None, "all"):
            return True
        if scope in ("law", "not-law"):
            return self.is_law_page(p) == (scope == "law")
        if scope == "sample":
            return "/assets/img/samples-" in p["html"]
        return p["lang"] == scope or p["lang2"] == scope

    def run_guards(self, p, text):
        for chk in self.rules["checks"]:
            if not self.in_scope(p, chk.get("scope")):
                continue
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

    def check_withdrawn_names(self, built):
        """data/rules.yaml withdrawn_names, in the guard text and every URL-like value of each page, in
        dist/ file names and in the redirect rules."""
        wn = self.rules.get("withdrawn_names") or {}
        lists = {k: set(wn.get(k) or []) for k in ("anywhere", "outside_law", "on_sample_pages")}
        for k, digests in lists.items():
            for x in digests:
                if not re.fullmatch(r"[0-9a-f]{64}", str(x)):
                    self.err(f"data/rules.yaml: withdrawn_names.{k}: {x!r} is not a sha256 digest")
        why = {"anywhere": "a name the owner withdrew on 27 September 2026 (OWNER_DECISIONS.md)",
               "outside_law": "a name that may appear only on the law pages",
               "on_sample_pages": "not on a page that shows the pipeline sample"}

        def scan(where, text, active):
            digests = set().union(*(lists[k] for k in active))
            if not digests:
                return
            for a, b, cand in name_candidates(text):
                if name_digest(cand) in digests:
                    kind = next(k for k in active if name_digest(cand) in lists[k])
                    self.err(f"{where}: [withdrawn-name] {why[kind]}: {cand!r} in “…{text.lower()[max(0, a - 40):b + 40]}…”")
                    return

        for p in built:
            active = ["anywhere"]
            if not self.is_law_page(p):
                active.append("outside_law")
            if self.in_scope(p, "sample"):
                active.append("on_sample_pages")
            h = re.sub(r"(?s)<style\b.*?</style>", " ", p["html"])
            values = re.findall(r'\b(?:href|src|srcset|id|action)="([^"]*)"', h)
            values += [v for v in re.findall(r'\bcontent="([^"]*)"', h) if v.startswith(("http", "/"))]
            for ld in re.findall(r'<script type="application/ld\+json">(.*?)</script>', h, re.S):
                values += re.findall(r'"((?:https?:)?/[^"]*)"', ld.replace("\\/", "/"))
            scan(p["url"], self.guard_text(p) + " ¶ " + " ¶ ".join(htmllib.unescape(v) for v in dict.fromkeys(values)), active)
        for f in sorted(self.dist.rglob("*")):
            if f.is_file():
                scan("dist", "/" + f.relative_to(self.dist).as_posix(), ["anywhere", "outside_law"])
        for src, dst, code in self.redirects:
            scan("data/redirects.yaml", f"{src} ¶ {dst}", ["anywhere", "outside_law"])

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

    @staticmethod
    def native_list(lang):
        """reviews.yaml pending list for a language: native_pt for pt-MZ (its first name), else native_<lang>."""
        return "native_pt" if lang == "pt-MZ" else "native_" + lang.lower().replace("-", "_")

    def pending_reviews(self):
        """data/reviews.yaml pending lists as {list: {url: (owner decision date, list name)}}. A list is a list
        of URLs or {url, owner_decision} items, a mapping url -> reviewer or {reviewer, owner_decision}, or
        {owner_decision, pages: <either>}; an entry's date wins over its list's, which wins over the file's."""
        rv = self.reviews
        top = rv.get("owner_decision")
        out = {}
        for name, val in (rv.get("pending") or {}).items():
            if not (name == "counsel" or name.startswith("counsel_") or name.startswith("native_")):
                self.err(f"data/reviews.yaml: pending.{name}: lists are native_<lang> or counsel / counsel_<name>")
                continue
            date, entries = top, val
            if isinstance(val, dict) and "pages" in val:
                date, entries = val.get("owner_decision", top), val.get("pages")
            items = {}
            if isinstance(entries, dict):
                for url, v in entries.items():
                    items[url] = v.get("owner_decision", date) if isinstance(v, dict) else date
            else:
                for e in entries or []:
                    if isinstance(e, dict):
                        items[e.get("url")] = e.get("owner_decision", date)
                    else:
                        items[e] = date
            for url, d in items.items():
                if not isinstance(url, str) or not url.startswith("/"):
                    self.err(f"data/reviews.yaml: pending.{name}: {url!r} is not a page URL")
                if not isinstance(d, dt.date):
                    self.err(f"data/reviews.yaml: pending.{name}: {url} has no owner_decision date (YYYY-MM-DD)")
            out[name] = {url: (d, name) for url, d in items.items()}
        return out

    def check_reviews(self, pages):
        """Native and counsel reviews (data/reviews.yaml): a page that needs one and has neither the sign-off
        in its front matter nor an entry on the owner's pending lists fails the build. Every page in a language
        other than English needs a native review; law pages and counsel_required pages need counsel."""
        rv = self.reviews
        pending = self.pending_reviews()
        counsel_pending = {}
        for name, items in pending.items():
            if name == "counsel" or name.startswith("counsel_"):
                counsel_pending.update(items)
        counsel_required = set(rv.get("counsel_required") or [])
        seen = set()
        for p in pages:
            if p["status"] != "published" or p.get("draft"):
                continue
            m, url = p["meta"], p["url"]
            needs = []
            if p["lang2"] != "en":
                name = self.native_list(p["lang"])
                what = "a native Mozambican review" if p["lang"] == "pt-MZ" else f"a native review ({p['lang']})"
                needs.append((name, "reviewed_on", "reviewed_by_role", pending.get(name, {}), what))
            if p["template"] == "law" or url in counsel_required or m.get("counsel_required"):
                needs.append(("counsel", "counsel_reviewed_on", "counsel_reviewed_by_role", counsel_pending,
                              "a counsel review"))
            for kind, on, by, pend, what in needs:
                seen.add((kind, url))
                if m.get(on):
                    if not m.get(by):
                        self.err(f"{url}: {on} is set without {by}")
                    if url in pend:
                        self.err(f"{url}: signed off ({on}) but still listed under pending.{pend[url][1]} in data/reviews.yaml")
                elif url in pend:
                    self.warn(f"{url}: published before {what} (owner decision {pend[url][0]}; "
                              f"data/reviews.yaml pending.{pend[url][1]})")
                else:
                    self.err(f"{url}: needs {what} before it is published: set {on} and {by} after a real "
                             f"sign-off, or keep it status: draft; only an owner decision to publish first puts it on "
                             f"pending.{kind} in data/reviews.yaml, with that decision's date")
        for name, items in pending.items():
            kind = "counsel" if name == "counsel" or name.startswith("counsel_") else name
            for url in sorted(items):
                if (kind, url) not in seen:
                    self.warn(f"data/reviews.yaml: pending.{name} lists {url}, which is not a published page that needs it")

    def check_draft_leaks(self, built):
        """Nothing published may point into a draft section: no file under its prefix, hreflang alternate,
        sitemap entry, country-menu item, country-sites button, /countries entry, JSON-LD reference, _redirects
        target, _headers rule, banner suggestion or internal link. A --drafts build instead checks that every
        draft page is noindex, shows the draft banner and stays out of sitemaps and hreflang."""
        draft = {lk for lk in self.locales if not self.is_live(lk)}
        live_cc = {l["country"] for lk, l in self.locales.items() if l.get("country") and self.is_live(lk)}
        draft_cc = {l["country"] for lk, l in self.locales.items() if l.get("country") and lk in draft} - live_cc
        draft_codes = {c for lk in draft for c in self.locales[lk]["hreflang"]}
        code_owner = {c: lk for lk in draft for c in self.locales[lk]["hreflang"]}
        draft_langs = {self.locales[lk]["lang"] for lk in draft}
        draft_names = {n for lk in draft if self.locales[lk]["country"] in draft_cc
                       for n in (self.locales[lk]["country_name"], (self.country_of(lk) or {}).get("name"))}
        found = []

        def in_draft(value):
            v = str(value or "").strip()
            if v.startswith(self.base + "/") or v == self.base:
                v = v[len(self.base):] or "/"
            if not v.startswith("/") or v.startswith("//"):
                return None
            lk = self.section_of_path(re.split(r"[?#]", v)[0])
            return lk if lk in draft else None

        def leak(where, what, value, lk):
            found.append(f"[draft-leak] {where}: {what} → {value} (draft section {lk})")

        # the geo banner never suggests a draft, in any build
        for f in sorted((self.dist / "assets/js").glob("region.*.js")):
            m = re.search(r"const DATA = (\{.*?\});\n", f.read_text(encoding="utf-8"), re.S)
            data = json.loads(m.group(1)) if m else {}
            for cc, site in (data.get("sites") or {}).items():
                if cc in draft_cc:
                    leak(f"/assets/js/{f.name}", "geo banner", cc, next(k for k in draft if self.locales[k]["country"] == cc))
                for code, home, _ in site.get("links", []):
                    if (lk := in_draft(home)) or code in draft_langs:
                        leak(f"/assets/js/{f.name}", "geo banner", home, lk or code)
        sitemap_locs = []
        for f in sorted(self.dist.glob("sitemap*.xml")):
            sitemap_locs += [(f.name, u) for u in re.findall(r"<loc>([^<]+)</loc>", f.read_text(encoding="utf-8"))]

        if self.drafts:
            locs = {u for _, u in sitemap_locs}
            for p in built:
                if not p.get("draft"):
                    continue
                if not re.search(r'<meta name="robots" content="noindex', p["html"]):
                    found.append(f"[draft-page] {p['url']}: a draft page must be noindex")
                if DRAFT_BANNER not in p["html"]:
                    found.append(f"[draft-page] {p['url']}: a draft page must show the “{DRAFT_BANNER}” banner")
                if p["abs_url"] in locs or p.get("alternates"):
                    found.append(f"[draft-page] {p['url']}: a draft page is never in a sitemap or an hreflang cluster")
            for msg in found:
                self.err(msg)
            return

        for f in sorted(self.dist.rglob("*")):
            if f.is_file() and (lk := in_draft("/" + f.relative_to(self.dist).as_posix())):
                leak("dist", "a file in a draft section", "/" + f.relative_to(self.dist).as_posix(), lk)
        for p in built:
            refs = PageRefs()
            refs.feed(p["html"])
            for what, value, code in refs.refs:
                lk = in_draft(value) or (code_owner.get(code) if code in draft_codes else None)
                if lk:
                    leak(p["url"], what, f"{code} {value}" if code else value, lk)
            for ld in refs.jsonld:
                try:
                    doc = json.loads(ld.replace("<\\/", "</"))
                except json.JSONDecodeError:
                    continue
                stack = [(None, doc)]
                while stack:
                    key, v = stack.pop()
                    if isinstance(v, dict):
                        stack += list(v.items())
                    elif isinstance(v, list):
                        stack += [(key, x) for x in v]
                    elif isinstance(v, str):
                        if lk := in_draft(v):
                            leak(p["url"], "JSON-LD reference", v, lk)
                        elif key == "inLanguage" and v in draft_langs:
                            leak(p["url"], "JSON-LD reference", f"inLanguage {v}",
                                 next(k for k in draft if self.locales[k]["lang"] == v))
                        elif key == "name" and v in draft_names:
                            leak(p["url"], "JSON-LD reference", f"country {v}",
                                 next(k for k in draft if v in (self.locales[k]["country_name"],
                                                                (self.country_of(k) or {}).get("name"))))
        draft_maps = {self.locales[lk]["sitemap"] for lk in draft} - {self.locales[lk]["sitemap"] for lk in self.locales
                                                                       if lk not in draft}
        for name, u in sitemap_locs:
            if lk := in_draft(u):
                leak(name, "sitemap entry", u, lk)
            elif u.rsplit("/", 1)[-1] in draft_maps:
                leak(name, "sitemap entry", u, next(k for k in draft if self.locales[k]["sitemap"] == u.rsplit("/", 1)[-1]))
        for src, dst, _ in self.redirects:
            if lk := in_draft(dst):
                self.warn(f"data/redirects.yaml: {src} -> {dst} is left out of _redirects while {lk} is a draft section")
        if (self.dist / "_redirects").exists():
            for line in (self.dist / "_redirects").read_text(encoding="utf-8").splitlines():
                parts = line.split()
                if len(parts) >= 2 and not line.startswith("#") and (lk := in_draft(parts[1])):
                    leak("_redirects", "_redirects target", f"{parts[0]} -> {parts[1]}", lk)
        if (self.dist / "_headers").exists():
            for line in (self.dist / "_headers").read_text(encoding="utf-8").splitlines():
                if line.startswith("/") and (lk := in_draft(line.strip().rstrip("*"))):
                    leak("_headers", "_headers rule", line.strip(), lk)
        for msg in dict.fromkeys(found):
            self.err(msg)

    def check_deploy_config(self):
        """Cloudflare Pages must serve dist/, not the repository root (site/ sources would be public)."""
        cfg = ROOT / "wrangler.toml"
        if self.dist.resolve() != (ROOT / "dist").resolve():
            return
        text = cfg.read_text(encoding="utf-8") if cfg.exists() else ""
        if not re.search(r'(?m)^pages_build_output_dir\s*=\s*"\./dist"\s*$', text):
            self.err('wrangler.toml at the repository root must set pages_build_output_dir = "./dist"')

    # ------------------------------------------------------------------ main
    def check_i18n(self):
        def keys(d, pre=""):
            out = set()
            for k, v in d.items():
                out.add(pre + k)
                if isinstance(v, dict) and pre + k != "foreign_pages":   # each language lists its own
                    out |= keys(v, pre + k + ".")
            return out

        sets = {n: keys(d) for n, d in self.i18n.items()}
        allk = set().union(*sets.values())
        optional = lambda k: any(k == o or k.startswith(o + ".") for o in I18N_FALLBACK)
        for n, s in sets.items():
            missing = sorted(allk - s)
            for k in missing:
                if n != "en" and optional(k):
                    if not optional(k.rsplit(".", 1)[0]) or k.rsplit(".", 1)[0] not in missing:
                        self.warn(f"data/i18n/{n}.yaml: no '{k}'; the English string is used")
                else:
                    self.err(f"data/i18n/{n}.yaml: missing key '{k}'")

    def run(self):
        if self.drafts and (why := dist_refusal(self.dist)):
            self.err(f"--drafts never writes into a checkout's dist/ ({why}); pass --dist <scratch directory>")
            return self.report([], [])
        self.check_locales()
        self.check_i18n()
        pages = self.read_pages()
        self.check_reviews(pages)
        self.check_deploy_config()
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
        sections = {"global"} | {p["loc_key"] for p in pages}
        self.navs = {lk: self.nav_for(lk) for lk in self.locales if lk in sections}
        self.footers = {lk: self.footer_for(lk) for lk in self.navs}
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
        self.check_withdrawn_names(built)
        self.check_duplicates(built)
        self.check_similarity(built)
        self.check_law()
        self.check_draft_leaks(built)
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
        # withdrawn names: stand-ins (dummy_names below), since the real ones are never spelled out here
        "withdrawn":      ("global", append("Our sample is the QX-07 line."), "[withdrawn-name]"),
        "withdrawn-pt":   ("mz-pt", append("A amostra fica em Heron–Crest."), "[withdrawn-name]"),
        "withdrawn-id":   ("global", append("## Earlier view {#qx7-route}\n\nText."), "[withdrawn-name]"),
        "withdrawn-dot":  ("global", append("Our sample is the QX.7 line."), "[withdrawn-name]"),
        "svg-text":       ("global", append('<svg viewBox="0 0 120 20" role="img"><text x="0" y="14">Real-time alerts</text></svg>'),
                           "[overstatement]"),
        "draft-link":     ("global", append("See [our Zambia site](/zm/)."), "[draft-leak]"),
        "law-only-name":  ("global", append("Our sample is near the Morlock field."), "[withdrawn-name]"),
        "on-sample-page": ("global/results.md", append("See regulation 12/3456."), "[withdrawn-name]"),
        "law-page-ok":    ("mz/en/50m-protection-zone.md", append("The Morlock corridor has its own zone."), None),
        "other-page-ok":  ("global/faq.md", append("See regulation 12/3456."), None),
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
    dummy_names = {"anywhere": ["QX-7", "Heron Crest"], "outside_law": ["Morlock"], "on_sample_pages": ["12/3456"]}

    def make(**kw):
        b = Build(quiet=True, **kw)
        wn = b.rules.setdefault("withdrawn_names", {})
        for k, names in dummy_names.items():
            wn[k] = list(wn.get(k) or []) + [name_digest(n) for n in names]
        return b

    (CACHE_DIR / "selftest").mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="run-", dir=CACHE_DIR / "selftest"))
    failed = []
    try:
        clean = make(dist=tmp / "clean" / "dist")
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
            elif lk.endswith(".md"):
                target = content / lk
            else:
                target = content / "global/faq.md"
            target.write_text(mutate(target.read_text(encoding="utf-8")), encoding="utf-8")
            b = make(content_dir=content, dist=tmp / name / "dist")
            code = b.run()
            if expect is None:
                ok, hit = code == 0, [f"builds clean ({len(b.errors)} errors)"]
            else:
                hit = [e for e in b.errors if expect in e]
                ok = code == 1 and hit
            print(f"{'PASS' if ok else 'FAIL'}  {name:15s} {(hit or b.errors or ['no error raised'])[0][:105]}")
            if not ok:
                failed.append(name)
        # redirects: a withdrawn name in a rule, and a splat that would hide the published sample images
        for name, rule, expect in (("redirect-name", ["/assets/img/samples-qx7-*", "/results", 301], "[withdrawn-name]"),
                                   ("redirect-hides", ["/assets/img/samples-s*", "/results", 301], "would hide")):
            b = make(dist=tmp / name / "dist")
            b.redirects = b.redirects + [rule]
            code = b.run()
            hit = [e for e in b.errors if expect in e]
            print(f"{'PASS' if code == 1 and hit else 'FAIL'}  {name:15s} {(hit or b.errors or ['no error raised'])[0][:105]}")
            if not (code == 1 and hit):
                failed.append(name)
        # near-duplicate: an English country page that copies the global one
        content = tmp / "dup" / "content"
        shutil.copytree(SITE / "content", content)
        shutil.copytree(SITE / "tests/fixtures/za-dup", content / "za", dirs_exist_ok=True)
        # a whole-page copy (front matter included, so related cards and buttons match too) with a new title
        meta, body = split_front_matter((content / "global/results.md").read_text(encoding="utf-8"), "results")
        meta.update(title="Sample outputs South Africa | AfriScan", h1="Sample outputs in South Africa",
                    description="A copy of the global sample page with a different title, which the build must refuse as a near duplicate.")
        meta.pop("nav_group", None)
        (content / "za/results.md").write_text(
            "---\n" + yaml.safe_dump(meta, allow_unicode=True, sort_keys=False) + "---\n" + body, encoding="utf-8")
        b = make(content_dir=content, dist=tmp / "dup" / "dist")
        b.run()
        hit = [e for e in b.errors if "near-duplicate" in e]
        print(f"{'PASS' if hit else 'FAIL'}  {'near-duplicate':15s} {hit[0][:105] if hit else 'no error raised'}")
        if not hit:
            failed.append("near-duplicate")
        # whole new pages: a Portuguese page nobody has reviewed, a law page without counsel review, and a
        # social-card headline too long to fit (the card must never drop words)
        new_pages = {
            "pt-unreviewed": ("mz/pt/pagina-nova.md", "---\ntitle: Página nova de teste | AfriScan\ndescription: Uma página portuguesa nova, sem revisão de um falante nativo, que a construção tem de recusar.\nh1: Página nova\n---\nTexto.\n", "needs a native Mozambican review"),
            "law-unreviewed": ("za/drone-rules-copy.md", "---\ntitle: Drone Rules Copy for Testing | AfriScan\ndescription: A second South African law page with no counsel review, which the build must refuse to publish.\nh1: Drone rules copy\ntemplate: law\nlaw: za\n---\nText.\n", "needs a counsel review"),
            "og-overflow": ("global/og-test.md", "---\ntitle: Social Card Overflow Test | AfriScan\ndescription: A page whose social-card headline is far too long to fit, which the build must refuse rather than cut.\nh1: Card test\nog:\n  headline: " + "A very long headline that goes on and on " * 4 + "\n---\nText.\n", "social card headline does not fit"),
        }
        for name, (rel, text, expect) in new_pages.items():
            content = tmp / name / "content"
            shutil.copytree(SITE / "content", content)
            (content / rel).write_text(text, encoding="utf-8")
            b = make(content_dir=content, dist=tmp / name / "dist")
            code = b.run()
            hit = [e for e in b.errors if expect in e]
            print(f"{'PASS' if code == 1 and hit else 'FAIL'}  {name:15s} {(hit or b.errors or ['no error raised'])[0][:105]}")
            if not (code == 1 and hit):
                failed.append(name)
        failed += selftest_drafts(make, tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print("selftest:", "all guards fired" if not failed else f"FAILED {failed}")
    return 1 if failed else 0


def selftest_drafts(make, tmp):
    """Draft sections (data/locales.yaml status: draft). A normal build never builds them and nothing it
    publishes may point into one; a --drafts build marks their pages and never writes into a dist/.
    Each seeded leak stands for a regression in one place that could publish a pointer to a draft."""
    draft = next((lk for lk, l in load_yaml(SITE / "data/locales.yaml").items() if l.get("status") == "draft"), None)
    if not draft:
        print("SKIP  draft sections: every section in data/locales.yaml is live")
        return []
    loc = load_yaml(SITE / "data/locales.yaml")[draft]
    home, code = loc["prefix"] + "/", loc["hreflang"][0]
    fixture = SITE / "tests/fixtures/demo/content" / loc["content"] / "index.md"

    def with_home(content):
        (content / loc["content"]).mkdir(parents=True, exist_ok=True)
        shutil.copy2(fixture, content / loc["content"] / "index.md")

    def wrap(b, method, after):
        orig = getattr(b, method)
        setattr(b, method, lambda *a, **k: after(orig(*a, **k), *a))

    def add_item(result, item):
        if result.get("groups"):
            result["groups"][0]["items"] = result["groups"][0]["items"] + [item]
        return result

    def append_to(b, name, text):
        f = b.dist / name
        f.write_text(f.read_text(encoding="utf-8").replace("</urlset>", text + "</urlset>") if name.endswith(".xml")
                     else f.read_text(encoding="utf-8") + text, encoding="utf-8")

    def strip_draft_marks(b, what):
        def after(_, p):
            if p.get("draft"):
                p["html"] = (p["html"].replace(DRAFT_BANNER, "") if what == "banner" else
                             p["html"].replace('content="noindex, nofollow"', 'content="index, follow"'))
        wrap(b, "render_page", after)

    item = {"section": draft, "key": draft, "label": loc["label"], "title": loc["label"], "href": home, "lang": code,
            "text": "", "langs": "", "links": [], "current": False, "draft": False}
    fake = tmp / "fake-checkout"
    (fake / "site").mkdir(parents=True, exist_ok=True)
    (fake / "site/build.py").write_text("", encoding="utf-8")
    # name: (drafts, dist, add the draft home, patch, expected error or None for a clean build)
    cases = {
        "draft-home-ok":   (False, None, True, None, None),
        "drafts-empty-ok": (True, None, False, None, None),
        "drafts-ok":       (True, None, True, None, None),
        "drafts-in-dist":  (True, ROOT / "dist", False, None, "--drafts never writes"),
        "drafts-in-other": (True, fake / "dist/preview", False, None, "--drafts never writes"),
        "draft-built":     (False, None, True, lambda b: setattr(b, "section_built", lambda lk: True),
                            "a file in a draft section"),
        "draft-hreflang":  (False, None, False, lambda b: wrap(b, "alternates", lambda r, p: r and r + [
                                {"code": code, "href": b.base + home}]), "hreflang alternate"),
        "draft-menu":      (False, None, False, lambda b: wrap(b, "region_links", lambda r, p: r + [item]),
                            "region-menu item"),
        "draft-sites":     (False, None, False, lambda b: wrap(b, "cmp_country_sites", lambda r, *a: add_item(r, item)),
                            "country-sites button"),
        "draft-countries": (False, None, False, lambda b: wrap(b, "cmp_countries", lambda r, *a: add_item(r, item)),
                            "/countries entry"),
        "draft-jsonld":    (False, None, False, lambda b: wrap(b, "served_countries", lambda r: r + [
                                (b.country_of(draft) or {}).get("name", loc["country_name"])]), "JSON-LD reference"),
        "draft-sitemap":   (False, None, False, lambda b: wrap(b, "write_sitemaps", lambda r: append_to(
                                b, "sitemap-global.xml", f"  <url>\n    <loc>{b.base}{home}</loc>\n  </url>\n")),
                            "sitemap entry"),
        "draft-redirect":  (False, None, False, lambda b: wrap(b, "write_site_files", lambda r: append_to(
                                b, "_redirects", f"/old-{draft} {home} 301\n")), "_redirects target"),
        "draft-headers":   (False, None, False, lambda b: wrap(b, "write_site_files", lambda r: append_to(
                                b, "_headers", f"\n{loc['prefix']}/*\n  Content-Language: {loc['lang']}\n")),
                            "_headers rule"),
        "draft-banner":    (True, None, True, lambda b: setattr(b, "banner_sections", lambda: [
                                k for k in b.live_locales if b.locales[k]["country"]]), "geo banner"),
        "draft-unmarked":  (True, None, True, lambda b: strip_draft_marks(b, "banner"), "“Draft: not published” banner"),
        "draft-indexed":   (True, None, True, lambda b: strip_draft_marks(b, "robots"), "must be noindex"),
    }
    failed = []
    for name, (drafts, dist, add_home, patch, expect) in cases.items():
        content = tmp / name / "content"
        shutil.copytree(SITE / "content", content)
        if add_home:
            with_home(content)
        b = make(content_dir=content, dist=dist or tmp / name / "dist", drafts=drafts)
        if patch:
            patch(b)
        code_ = b.run()
        if expect is None:
            sec_dir = b.dist / loc["prefix"].lstrip("/")
            stray = [] if drafts or not sec_dir.exists() else list(sec_dir.rglob("*"))
            ok = code_ == 0 and not stray
            hit = [f"builds clean ({len(b.errors)} errors" + (f", {len(stray)} files under {home}" if stray else "") + ")"]
        else:
            hit = [e for e in b.errors if expect in e]
            ok = code_ == 1 and hit
        print(f"{'PASS' if ok else 'FAIL'}  {name:15s} {(hit or b.errors or ['no error raised'])[0][:105]}")
        if not ok:
            failed.append(name)
            for e in b.errors[:5]:
                print("      ", e)
    return failed


def demo(out):
    """Build the real content plus the template fixtures in tests/fixtures/demo into OUT (never deploy).
    Fixtures only fill gaps: a real page or law file with the same path always wins, so the demo keeps
    building once a country section has its own pages. It is a drafts build, so the draft sections'
    fixture homes fill the country selector (every country in data/locales.yaml)."""
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
    return Build(content_dir=work, dist=Path(out), law_dir=law, drafts=True).run()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--drafts", action="store_true", help="also build status: draft pages (never deploy this)")
    ap.add_argument("--selftest", action="store_true", help="check that every guard fails the build")
    ap.add_argument("--name-digest", metavar="NAME", help="print the digest data/rules.yaml withdrawn_names stores for NAME")
    ap.add_argument("--dist", default=str(ROOT / "dist"), help="output directory (default: dist/; --drafts needs another one)")
    ap.add_argument("--demo", metavar="OUT", help="build content + template fixtures into OUT, for checking templates")
    args = ap.parse_args()
    if args.name_digest:
        print(name_digest(args.name_digest))
        return 0
    if args.selftest:
        return selftest()
    if args.demo:
        return demo(args.demo)
    return Build(dist=Path(args.dist), drafts=args.drafts).run()


if __name__ == "__main__":
    sys.exit(main())
