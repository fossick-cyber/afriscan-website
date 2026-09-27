"""Sample galleries (:::sample-gallery) and the guards on sample images.

A gallery is drawn from a data/samples/<name>.json file written by tools/make_samples.py: an overview
and close-up views, each with an image per language, its caption, count text, credit and alt text.

Guards (run by build.py on every build):
  google-imagery  Google-imagery masters live in site/images/samples/google/. Each one must be declared
                  by a data/samples/*.json file whose imagery.provider is "google", with an attribution
                  for its language that is drawn on the image itself (drawn_text) and repeated in its
                  credit. Masters carry no EXIF, XMP, ICC profile or comment. On a page, every Google
                  image must sit inside a <figure> whose text carries the attribution in the page's
                  language, and must be the version drawn in the page's language.
  image text      The text drawn on every sample image (drawn_text and image_text in the data files)
                  goes through the page guards (pricing, overstatement, pt-MZ vocabulary…) for each page
                  that shows the image, and through the withdrawn-names guard for every image.
"""
import html as htmllib
import re
import shutil
from pathlib import Path

from PIL import Image
from markupsafe import Markup

from lib.content import ContentError

SITE = Path(__file__).resolve().parent.parent
GOOGLE_DIR = "samples/google"
SPEC = {"required": {"data"}, "allowed": {"views", "overview", "cols", "size", "priority", "legend"}}
IMG_URL = re.compile(r"/assets/img/([a-z0-9-]+?)-[0-9a-f]{8}-\d+\.(?:avif|webp|jpg)")
CLEAN_INFO = {"jfif", "jfif_version", "jfif_unit", "jfif_density", "progressive", "progression", "dpi"}
SIZES = {"wide": "(max-width: 1188px) 100vw, 1140px", "half": "(max-width: 900px) 100vw, 560px",
         "2": "(max-width: 640px) 100vw, (max-width: 1188px) 50vw, 561px"}


def _items(d):
    return list(d.get("views") or []) + ([d["overview"]] if d.get("overview") else [])


def google_images(samples):
    """{image master name: facts} for every image a data file declares as Google imagery."""
    out = {}
    for stem, d in samples.items():
        imagery = d.get("imagery")
        if not isinstance(imagery, dict) or imagery.get("provider") != "google":
            continue
        attribution = imagery.get("attribution") or {}
        for item in _items(d):
            for lang, src in (item.get("src") or {}).items():
                out[src] = {"data": stem, "lang": lang, "id": item.get("id"), "attribution": attribution.get(lang),
                            "attributions": attribution, "drawn": (item.get("drawn_text") or {}).get(lang) or [],
                            "credit": (item.get("credit") or {}).get(lang) or ""}
    return out


def image_texts(samples):
    """{dist file stem: [every string drawn on the image]} across all data/samples files."""
    out = {}
    for d in samples.values():
        for name, texts in (d.get("image_text") or {}).items():
            out[name.replace("/", "-")] = list(texts)
        for item in _items(d):
            for lang, src in (item.get("src") or {}).items():
                out[src.replace("/", "-")] = list((item.get("drawn_text") or {}).get(lang) or [])
    return out


def context(build, p, b, ctx):
    """Template context for :::sample-gallery{data="…" views="A,C" overview="true|false" cols="1|2" …}."""
    where, a = ctx["where"], b.attrs
    data = build.samples.get(a["data"])
    if not data or not data.get("views"):
        raise ContentError(f"{where}: no gallery in data/samples/{a['data']}.json (run site/tools/make_samples.py)")
    lang = p["lang2"]
    strings = (data.get("strings") or {}).get(lang)
    if not strings:
        raise ContentError(f"{where}: data/samples/{a['data']}.json has no '{lang}' strings; add the language to "
                           f"tools/make_samples.py and rerun it")
    imagery = data.get("imagery") or {}
    if imagery.get("provider") == "google" and not (imagery.get("attribution") or {}).get(lang):
        raise ContentError(f"{where}: data/samples/{a['data']}.json has no '{lang}' Google attribution")
    for k, allowed in (("overview", ("true", "false")), ("cols", ("1", "2")), ("size", ("wide", "half")),
                       ("priority", ("true", "false")), ("legend", ("true", "false"))):
        if a.get(k) and a[k] not in allowed:
            raise ContentError(f"{where}: ':::sample-gallery' {k} must be one of {list(allowed)}")
    by_id = {v["id"]: v for v in data["views"]}
    ids = [x.strip() for x in a.get("views", "").split(",") if x.strip()] or list(by_id)
    bad = [i for i in ids if i not in by_id]
    if bad:
        raise ContentError(f"{where}: unknown view ids {bad} (data/samples/{a['data']}.json has {list(by_id)})")
    cols = a.get("cols", "2")
    size = a.get("size", "wide")
    if cols == "2" and a.get("size"):
        raise ContentError(f"{where}: ':::sample-gallery' size applies to cols=\"1\" only")
    priority = a.get("priority") == "true"
    show_overview = a.get("overview", "true") == "true" and data.get("overview")

    def fig(item, sizes, first):
        if lang not in (item.get("src") or {}):
            raise ContentError(f"{where}: view {item.get('id')} has no '{lang}' image")
        pic = build.images.picture(item["src"][lang], item["alt"][lang], sizes=sizes, priority=first and priority,
                                   page_url=p["url"])
        return {"picture": Markup(pic), "caption": item["caption"][lang], "text": (item.get("text") or {}).get(lang),
                "credit": item["credit"][lang], "id": item.get("id")}

    two = cols == "2" and len(ids) > 1
    overview = fig(data["overview"], SIZES["wide"], True) if show_overview else None
    views = [fig(by_id[i], SIZES["2" if two else size], not overview and n == 0) for n, i in enumerate(ids)]
    summary = None                  # one view already gives its own count in its figure text
    if len(ids) > 1 and data.get("within_100_total") and strings.get("summary"):
        shown = len({r for i in ids for r in by_id[i].get("within_100_ids") or []})
        summary = strings["summary"].format(shown=shown, total=data["within_100_total"])
    return {"g": {"overview": overview, "views": views, "two": two, "size": size, "s": strings, "summary": summary,
                  "legend": a.get("legend", "true") == "true"}}


def check(build, built, name_candidates, name_digest):
    samples = build.samples
    gi = google_images(samples)
    err = build.err
    # declarations: attribution drawn on the image and repeated in its credit
    for name, info in sorted(gi.items()):
        where = f"data/samples/{info['data']}.json view {info['id']} ({info['lang']})"
        if not name.startswith(GOOGLE_DIR + "/"):
            err(f"{where}: [google-imagery] {name}: Google-imagery masters belong in site/images/{GOOGLE_DIR}/")
        att = info["attribution"]
        if not att:
            err(f"{where}: [google-imagery] no '{info['lang']}' attribution in imagery.attribution")
            continue
        if att not in info["drawn"]:
            err(f"{where}: [google-imagery] the image itself does not carry {att!r} (drawn_text)")
        if att not in info["credit"]:
            err(f"{where}: [google-imagery] the credit does not carry {att!r}")
    # every master in the Google folder is declared, and none carries metadata
    root = getattr(build, "sample_images_root", None) or SITE / "images"      # the self-test points it elsewhere
    folder = root / GOOGLE_DIR
    for f in sorted(folder.glob("*")) if folder.exists() else []:
        name = f"{GOOGLE_DIR}/{f.stem}"
        if f.suffix.lower() not in (".jpg", ".jpeg", ".png", ".webp"):
            continue
        if name not in gi:
            err(f"site/images/{name}{f.suffix}: [google-imagery] not declared in any data/samples/*.json with "
                f"imagery.provider \"google\" and its attribution (rerun site/tools/make_samples.py)")
    for f in sorted((root / "samples").rglob("*")):
        if f.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp"):
            with Image.open(f) as im:
                extra = sorted(set(im.info) - CLEAN_INFO)
                if extra or len(im.getexif()):
                    err(f"site/images/{f.relative_to(root)}: [image-metadata] strip {extra or ['exif']} "
                        f"(tools/make_samples.py saves masters without metadata)")
    # pages: each Google image inside a figure that carries the page language's attribution
    stems = {name.replace("/", "-"): info for name, info in gi.items()}
    for p in built:
        h = p.get("html") or ""
        if "/assets/img/" + GOOGLE_DIR.replace("/", "-") + "-" not in h:
            continue
        lang = p["lang2"]
        rest = h
        for fig in re.findall(r"(?s)<figure\b.*?</figure>", h):
            shown = {m.group(1) for m in IMG_URL.finditer(fig)} & stems.keys()
            if not shown:
                continue
            text = htmllib.unescape(re.sub(r"<[^>]+>", " ", re.sub(r"(?s)<(script|style|svg)\b.*?</\1>", " ", fig)))
            text = " ".join(text.split())
            ok = True
            for st in sorted(shown):
                info = stems[st]
                att = info["attributions"].get(lang)
                if not att:
                    err(f"{p['url']}: [google-imagery] {st}: no '{lang}' attribution in data/samples/{info['data']}.json")
                    ok = False
                elif att not in text:
                    err(f"{p['url']}: [google-imagery] a figure shows Google imagery ({st}) without {att!r} in its "
                        f"caption or credit")
                    ok = False
                if info["lang"] != lang:
                    err(f"{p['url']}: [google-imagery] {st} is the '{info['lang']}' version; use the '{lang}' one")
                    ok = False
            if ok:
                rest = rest.replace(fig, " ")
        leftover = sorted({m.group(1) for m in IMG_URL.finditer(rest)} & stems.keys())
        for st in leftover:
            err(f"{p['url']}: [google-imagery] Google imagery ({st}) shown outside a figure that carries its attribution")
    # the text drawn on the images: page guards per page, withdrawn names everywhere
    texts = image_texts(samples)
    wn = build.rules.get("withdrawn_names") or {}
    lists = {k: {str(x) for x in (wn.get(k) or [])} for k in ("anywhere", "outside_law", "on_sample_pages")}

    def scan(where, text, active):
        digests = set().union(*(lists[k] for k in active))
        for a_, b_, cand in name_candidates(text):
            if name_digest(cand) in digests:
                err(f"{where}: [withdrawn-name] a withdrawn name drawn on an image: {cand!r}")
                return

    for st, t in sorted(texts.items()):
        scan(f"image {st} (data/samples)", " ¶ ".join(t), ["anywhere", "outside_law", "on_sample_pages"])
    for p in built:
        shown = sorted({m.group(1) for m in IMG_URL.finditer(p.get("html") or "")} & texts.keys())
        drawn = [x for st in shown for x in texts[st]]
        if drawn and (not p["noindex"] or p.get("special")):
            build.run_guards(dict(p, url=f"{p['url']} (text drawn on its images)"), " ¶ ".join(drawn))


def selftest(make, tmp, site):
    """Seeded mistakes for the gallery guards; returns the names that did not fail as expected."""
    failed = []
    fig = ':::figure{src="samples/google/pipeline-view-a" alt="A satellite view" caption="A view of the route"}\n:::'

    def run(name, expect, content_mut=None, data_mut=None, page="global/faq.md", images=None):
        content = tmp / name / "content"
        shutil.copytree(site / "content", content)
        if content_mut:
            f = content / page
            f.write_text(content_mut(f.read_text(encoding="utf-8")), encoding="utf-8")
        b = make(content_dir=content, dist=tmp / name / "dist")
        if images:
            b.sample_images_root = images(tmp / name / "images")
        if data_mut:
            data_mut(b.samples)
        code = b.run()
        hit = [e for e in b.errors if expect in e]
        ok = code == 1 and hit
        print(f"{'PASS' if ok else 'FAIL'}  {name:15s} {(hit or b.errors or ['no error raised'])[0][:105]}")
        if not ok:
            failed.append(name)

    g = "sample-pipeline-google"
    run("google-credit", "[google-imagery]", content_mut=lambda raw: raw.rstrip() + "\n\n" + fig + "\n")
    run("google-undecl", "not declared", data_mut=lambda s: s[g].update(views=s[g]["views"][1:]))
    run("google-drawn", "does not carry", data_mut=lambda s: s[g]["views"][0]["drawn_text"].update(en=["View A"]))
    run("image-name", "[withdrawn-name]",
        data_mut=lambda s: s[g]["views"][2]["drawn_text"]["en"].append("QX-7 line"))
    run("image-claim", "[overstatement]", data_mut=lambda s: s["sample-pipeline"]["image_text"][
        "samples/sample-pipeline-register-km5-6"].append("Live monitoring"))
    # the English version of a Google view on a Portuguese page, even with the Portuguese credit
    fig_en = (':::figure{src="samples/google/pipeline-view-a" alt="Uma vista de satélite" caption="Vista A" '
              'credit="Imagens © Google"}\n:::')
    run("google-lang", "use the 'pt' one", page="mz/pt/resultados-de-exemplo.md",
        content_mut=lambda raw: raw.rstrip() + "\n\n" + fig_en + "\n")

    def with_exif(root):
        shutil.copytree(site / "images/samples", root / "samples")
        im = Image.new("RGB", (8, 8))
        exif = Image.Exif()
        exif[0x0110] = "Seeded camera model"
        im.save(root / "samples/seeded-metadata.jpg", exif=exif)
        return root

    run("image-metadata", "[image-metadata]", images=with_exif)
    return failed
