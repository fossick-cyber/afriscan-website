# AfriScan website: content guide for writers

This guide is for anyone adding or editing pages on afri-scan.com and its country sites (`/mz/`, `/mz/pt/`, `/za/`, `/ng/`). It covers the file format, where each kind of fact lives, the components you can use, and the guards that fail the build.

**Read first:** the owner's rules in `/home/claude/fhc-notes/website/OWNER_DECISIONS.md`. They win over anything else, including this guide. The page briefs are in `/home/claude/fhc-notes/website/industries/INDUSTRIES.FINAL.md` (industries and solutions), `industries/services.FINAL.md` (the service catalogue and its copy rules) and `regional/{mz,za,ng}.FINAL.md` (country facts, drone law, vocabulary). The ground truth for what the product does is `industries/capabilities.md`.

---

## 1. Build, check, preview

Everything runs with the app's virtualenv Python (Jinja2 3.1, Pillow 12, markdown-it-py, PyYAML):

```bash
cd /home/claude/afriscan-site
/opt/favhousecheck/.venv/bin/python3 site/build.py            # build dist/ and run every guard; exit 1 on any error
/opt/favhousecheck/.venv/bin/python3 site/build.py --selftest # prove each guard still fails on a seeded mistake (~1 min)
/opt/favhousecheck/.venv/bin/python3 site/build.py --demo /tmp/x   # real content + template fixtures (checks layouts; never deploy)
/opt/favhousecheck/.venv/bin/python3 site/build.py --drafts --dist /tmp/y   # include status: draft pages (local preview only)
```

Preview (the local server does not map `/x` to `x.html` the way Cloudflare Pages does, so open `/x.html`):

```bash
python3 -m http.server 5091 --directory /home/claude/afriscan-site/dist --bind 127.0.0.1
/opt/favhousecheck/.venv/bin/python3 site/tools/shoot.py http://127.0.0.1:5091 /tmp/shots / /results /za/   # 1440 + 390 screenshots, flags horizontal overflow
```

Read the screenshots. A page is not done until it looks right at 1440 px and 390 px.

A build prints every `ERROR` (build fails) and `WARN` (build passes; read it anyway), then a one-line summary. `dist/` is generated: never edit it by hand, and commit it together with the source change that produced it.

---

## 2. Where things live

```
site/
  build.py            generator and guards (this is the only program)
  lib/                content parsing, responsive images, social cards and icons
  content/            ONE FILE PER PAGE, Markdown with YAML front matter
    global/           afri-scan.com/…            (en-GB)
    mz/en/            afri-scan.com/mz/…         (en-MZ)
    mz/pt/            afri-scan.com/mz/pt/…      (pt-MZ)
    za/               afri-scan.com/za/…         (en-ZA)
    ng/               afri-scan.com/ng/…         (en-NG)
  data/
    site.yaml         brand, organisation JSON-LD, FormSubmit endpoint (do not change the endpoint)
    locales.yaml      the five sections: prefix, language, hreflang codes, selector labels
    catalogue.yaml    industries, solutions (U-codes) and services (S01–S45): names, blurbs, keys
    law/<cc>.yaml     drone-law instruments per country (mz, za, ng): the Sources tables come from here
    samples/t9.json   the T-9 sample register and segment ratings (written by tools/make_samples.py)
    i18n/en.yaml, i18n/pt-MZ.yaml   every interface string (nav, buttons, form, footer, 404…)
    glossary/pt-MZ.yaml             Mozambican Portuguese: banned Brazilian forms, AO90 warnings, preferred terms
    rules.yaml        the guard patterns
    redirects.yaml    path redirects -> dist/_redirects
    icons.yaml        inline SVG icons by name
  templates/          Jinja2 page templates, partials and components
  static/assets/      CSS, JS, font (copied; CSS and JS get content-hashed file names)
  images/             image masters; the build makes AVIF, WebP and JPEG at several widths
  tests/fixtures/     self-test and demo fixtures (never published)
  tools/              make_samples.py (T-9 images), shoot.py (screenshots)
functions/geo.js      the one Pages Function: /geo returns the visitor's country for the banner
dist/                 GENERATED site that Cloudflare Pages serves
```

---

## 3. Adding a page

1. Create a `.md` file in the right section folder. The path gives the URL:

   | File | URL |
   |---|---|
   | `content/global/index.md` | `/` |
   | `content/global/how-we-work.md` | `/how-we-work` |
   | `content/global/industries/oil-gas.md` | `/industries/oil-gas` |
   | `content/za/index.md` | `/za/` |
   | `content/za/pipelines.md` | `/za/pipelines` |
   | `content/mz/pt/lei-de-drones.md` | `/mz/pt/lei-de-drones` |

   Section homes have a trailing slash (`/za/`); every other page has none (`/za/pipelines`). Only a section home may be called `index.md`. Slugs are lowercase ASCII words joined by `-` (no accents: `zona-de-proteccao-parcial-50-metros`). `pt` is reserved under `/mz/`. You can override the path with `slug:`.

2. Write the front matter (between `---` lines) and the body (Markdown plus the components in §6).

3. Build. Fix every ERROR. Read every WARN.

4. Screenshot at 1440 and 390 and read the screenshots.

5. Commit the source and `dist/` together.

**A published slug is permanent.** If you rename or remove a page, add a line to `data/redirects.yaml` in the same commit.

### 3.1 A complete example

```markdown
---
key: oil-gas
template: industry
title: Pipeline Servitude & Right-of-Way Monitoring SA | AfriScan
description: Structures inside gas and liquid-fuel pipeline servitudes in South Africa, with distances to the line and change between surveys.
h1: Pipeline servitude monitoring in South Africa
crumb: Pipelines
eyebrow: Oil & gas
lead: A dated register of the structures inside each servitude width you give us, reviewed by a person.
buttons:
  - {label: Request a proposal, intent: proposal}
  - {label: See sample outputs, key: results}
related: [route-site-selection, change-detection, drone-surveys]
faq:
  - q: Do satellite surveys need a drone permit?
    a: No drone flies in satellite-based work.
---

::::section{id="servitudes" eyebrow="The problem" title="What stands inside the servitude"}
Body text in Markdown. Link to [the methodology](/methodology) or to a cluster key: [route selection](key:route-site-selection).

:::cards{cols="3"}
:::card{title="Structures by band" icon="corridor"}
Card text.
:::
:::
::::
```

### 3.2 Front-matter fields

Required on every page: `title`, `description`, `h1`.

| Field | What it does |
|---|---|
| `title` | `<title>` and og:title. ≤ 60 characters is the target, 65 the hard limit. End with ` \| AfriScan`. Unique across the site. |
| `description` | Meta description and og:description. 70–160 characters (170 hard limit). Unique across the site. |
| `h1` | The page's only H1. |
| `key` | **Cluster key for hreflang.** Pages in different sections with the same key are declared equivalent (see §4). Industry and solution pages must use their catalogue key. Default: `home` for a section home, otherwise a key unique to the page (no hreflang). |
| `template` | `home`, `page` (default), `industry`, `solution`, `country_home`, `law`, `law_hub`, `article`, `hub`, `contact`. See §5. |
| `status` | `published` (default) or `draft`. Drafts are never built unless you pass `--drafts`, so they cannot leak through hreflang or menus. |
| `slug` | Override the URL path within the section. |
| `crumb` | Short label for breadcrumbs and cards ("Pipelines"). Falls back to `nav_label`, then the catalogue name, then `h1`. |
| `eyebrow` | Small caps line above the H1. |
| `lead` | Paragraph under the H1 (inline Markdown). |
| `buttons` | Up to two buttons in the page head: `{label, intent}` (contact form with that intent), `{label, key}` (page with that cluster key, `anchor: "#id"` optional) or `{label, href}`. |
| `hero` | `home` and `country_home` only: `image` (master in `site/images/`, no extension), `alt`, `credit` (inline Markdown), `chips` (list of `{title, text}`). |
| `section` | Which hub this page sits under for breadcrumbs and `:::pages` lists: `industries`, `solutions`, `how`, `resources`, `countries`, `insights`. Industry, solution and article templates set it automatically. |
| `hub` | Marks the page as the hub for a section (one per section and language): the target of the top-level menu item and a breadcrumb level. |
| `parent` | Key of an extra breadcrumb parent. |
| `nav_group` | Puts the page in a header menu: `how`, `countries`, `resources` (industries and solutions come from the catalogue automatically). Country-only pages can also join `industries` or `solutions`. |
| `nav_order`, `nav_label`, `nav_blurb` | Menu order (low first), label and one-line description. |
| `summary`, `icon` | Card text and icon when this page appears in a `:::pages` list. |
| `related` | Keys shown as "Also useful" cards at the end. Catalogue keys without a page yet are skipped quietly; other unknown keys fail the build. |
| `used_in` | Solution pages: industry keys shown as "Used in:" under the H1. |
| `faq` | List of `{q, a}` (answer in Markdown). Rendered as an FAQ section and FAQPage JSON-LD. Keep answers true, short and local. |
| `faq_title`, `related_title` | Override those section headings. |
| `cta` | Closing band: `false` to remove it, or `{title, text, button, intent, href, secondary, secondary_href}`. Defaults come from i18n. |
| `service` | JSON-LD Service node: `{name, type, description}`, or `false`. Industry, solution and country-home pages get one automatically. Never add offers or prices. |
| `og` | Social card: `{headline, subline, alt}`. Defaults: the H1 and the "a person reviews every result" line. |
| `law`, `as_of` | Law pages: the country code of `data/law/<cc>.yaml`, and the "as of" date (default `law_as_of` in `site.yaml`, 26 September 2026). |
| `published`, `updated` | Articles: ISO dates. |
| `reviewed_on`, `reviewed_by_role` | pt-MZ pages: set ONLY when a native Mozambican reviewer has actually signed the page off. Until then the build warns; never invent a review date. |
| `noindex`, `sitemap` | `noindex: true` keeps a page out of search and the sitemap; `sitemap: false` keeps an indexable page out of the sitemap. |

**YAML gotcha:** any value that contains `: ` (colon space), or starts with a quote, `[`, `{`, `*`, `&` or `#`, must be in double quotes: `title: "Where We Work: Mozambique | AfriScan"`.

---

## 4. Clusters, hreflang and the country selector

A **cluster** is a set of pages that answer the same question for different audiences. The build writes reciprocal `hreflang` links for every published member of a cluster (codes from `locales.yaml`: `en` + `x-default` on the global page, `en-MZ`, `pt-MZ` + `pt`, `en-ZA`, `en-NG`), a self canonical on every page, and points the country selector at the equivalent page in each section (or that section's home when there is none). Clusters shrink automatically when a member is a draft.

Rules:
- **Cluster only true equivalents.** A South African drone-law page and a Nigerian one answer different questions: give them different keys. Mozambique's EN and PT versions of a country-only page share a key (for example `mz-drone-law`) and get `x-default` to the English one.
- **One page per section per key.** The build fails otherwise.
- **Never canonicalise a country page to the global page.** The build always writes a self canonical.
- **Same-language cluster members must not be near-duplicates** (§9).

Suggested keys and URLs. Industry and solution keys must match `data/catalogue.yaml`; country slugs follow `regional/architecture.FINAL.md` §1.6 and the country briefs.

| Key | Global | `/mz/` (EN) | `/mz/pt/` | `/za/` | `/ng/` |
|---|---|---|---|---|---|
| `home` | `/` | `/mz/` | `/mz/pt/` | `/za/` | `/ng/` |
| `contact` | `/contact` | `/mz/contact` | `/mz/pt/contacto` | `/za/contact` | `/ng/contact` |
| `how-we-work` | `/how-we-work` | | `/mz/pt/como-trabalhamos` | | |
| `oil-gas` | `/industries/oil-gas` | `/mz/pipelines` | `/mz/pt/gasodutos-e-oleodutos` | `/za/pipelines` | `/ng/oil-gas-pipelines` |
| `power-utilities` | `/industries/power-utilities` | `/mz/power-lines` | `/mz/pt/linhas-de-transporte-de-energia` | `/za/power-lines` | `/ng/power-transmission` |
| `mining` | `/industries/mining` | | | `/za/mining` | |
| `project-finance-esia` | `/industries/project-finance-esia` | | | | `/ng/esia-support` |
| `rail-roads` | `/industries/rail-roads` | | | `/za/rail-and-roads` | |
| `renewables`, `agriculture-forestry-nature`, `government` | `/industries/<key>` | | | | |
| `right-of-way-monitoring` | `/solutions/right-of-way-monitoring` | `/mz/protection-zone-survey` | `/mz/pt/levantamento-de-ocupacoes` | | |
| `resettlement-cut-off-baselines` | `/solutions/resettlement-cut-off-baselines` | `/mz/resettlement-cut-off-date` | `/mz/pt/reassentamento-data-de-corte` | | `/ng/resettlement-compensation-baselines` |
| `change-detection` | `/solutions/change-detection` | `/mz/repeat-surveys` | `/mz/pt/monitoria-periodica` | | |
| `evidence-packs` | `/solutions/evidence-packs` | | | never on `/za/` | `/ng/incident-evidence` |
| other solutions | `/solutions/<key>` (keys in catalogue.yaml) | | | | |
| `mz-drone-law` | | `/mz/drone-regulations` | `/mz/pt/lei-de-drones` | | |
| `za-drone-law`, `ng-drone-law` | | | | `/za/drone-regulations` | `/ng/drone-regulations` |
| `drone-regulations` | `/drone-regulations` (template `law_hub`) | | | | |

Country-only pages (procurement, POPIA, NDPA, 50 m protection zone, etc.) take their own key and no cluster, except Mozambique's EN↔PT pairs.

**The country selector and the banner** appear automatically as soon as a second section has a home page. The banner (bottom of the screen, never a redirect) suggests the visitor's country site from `/geo`, with a time-zone fallback, and only offers sections that exist.

---

## 5. Page templates

| Template | Use for | Notes |
|---|---|---|
| `home` | The global home | Full-bleed `hero` image with chips and credit. |
| `country_home` | `/mz/`, `/mz/pt/`, `/za/`, `/ng/` | Same hero pattern; the eyebrow defaults to the country name. Gets a Service node with `areaServed` = the country. |
| `industry` | Industry pages, global and country | Section defaults to `industries`. Service node. Brief: INDUSTRIES.FINAL §3–4 (hero, problem, sections with stable anchors, country context, how it works, honest-scope box, FAQ, CTA; 1,500–2,000 words global, 1,200–1,800 country). |
| `solution` | Solution pages | Section `solutions`; `used_in` renders "Used in:" under the H1. Service node. Brief: INDUSTRIES.FINAL §5 (800–1,400 words). |
| `law` | Country drone-law and compliance pages | Needs `law: <cc>`. Shows "As of 26 September 2026" and the not-legal-advice line in the head, and a Sources table built from `data/law/<cc>.yaml` at the end. JSON-LD carries `lastReviewed` and each instrument as a `Legislation` citation. |
| `law_hub` | `/drone-regulations` | Use `:::law-table` to render one row per country that has a law page. |
| `article` | Insights and country guides | Needs `published`. Shows the date and reading time; JSON-LD Article. Put them under `insights/` (global) or a country folder. |
| `hub` | Section hubs | Use `:::industries`, `:::solutions`, `:::pages` to list children. |
| `contact` | `/contact`, `/mz/pt/contacto`, `/za/contact`… | The form is built from i18n strings; the body (optional) goes in the sidebar. Keep `cta: false`. The FormSubmit endpoint lives in `site.yaml` and must not change. The thank-you page is `/thanks` (English) or `/mz/pt/obrigado`. |
| `page` | Anything else | |

404 and thank-you pages are generated from i18n strings (English everywhere; Portuguese under `/mz/pt/` once it has a home). They are noindex and never in the sitemap.

---

## 6. Components

Components are fenced blocks: an opening line `:::name{attr="value" …}` and a closing line of the same number of colons. A bare fence closes the innermost open block; using more colons for outer blocks (`::::section`, `:::cards`, `:::card`) keeps files readable. Markdown works inside every component. Unknown components or attributes fail the build with the line number.

Body content outside a `section` is wrapped in a plain white section automatically, so a simple page can be plain Markdown.

| Component | Attributes (* required) | Renders |
|---|---|---|
| `section` | `id`, `tone` (`light` default, `alt`, `dark`, `brand`), `eyebrow`, `title`, `lead`, `width` (`prose`), `class` | A full-width band. Alternate `light` and `alt`; use `dark` sparingly (one per page). |
| `cards` | `cols` (`2`, `3` default, `4`), `style` (`dark`, `plain`) | A responsive grid of the `card`, `figure` or other blocks inside it. |
| `card` | `title`*, `icon`, `eyebrow`, `tag`, `href` or `key`, `cta` | A card; with `href`/`key` the whole card is a link. |
| `steps` | `style` (`list` for a vertical list) | Numbered steps; put `step` blocks inside. |
| `step` | `title`*, `icon` | One step. |
| `callout` | `tone` (`note`, `scope`, `warn`, `legal`), `title`, `icon` | A boxed note. Use `scope` for "what we do and don't do", `legal` for law notes. |
| `columns` | `split` (`1-1`, `2-1`, `1-2`), `align` (`center`) | Two columns (stack below 900 px); put two `col` blocks inside. |
| `col` | `class` | One column. |
| `figure` | `src`*, `alt`*, `caption`, `credit`, `badge`, `size` (`wide` default, `full`, `narrow`, `half`, `third`), `priority` (`true` for the first large image only), `sizes` | A responsive picture (AVIF/WebP/JPEG, width and height set). Body text inside becomes the caption text. |
| `facts` | `cols` (`2`, `3`, `4`) | A grid of key facts from lines `- Term: value`. |
| `chips` | | Wrap a Markdown list to render it as pills. |
| `checklist` | `tone` (`no` for crosses) | Wrap a Markdown list: ticks or crosses. |
| `cta` | `title`, `text`, `button`, `intent` or `href`, `secondary`, `secondary_href` | An inline call to action box. |
| `industries` | `keys` (comma list), `tier` (`1`, `2`), `cols` | Catalogue industry cards, linked to each industry's page in this section (else the global page). |
| `solutions` | `keys`, `group` (`protect`, `baselines`, `change`, `capture`), `cols` | Catalogue solution cards, linked the same way. |
| `catalogue` | `services` (e.g. `S01,S08`), `groups` | The service list with public lines from `catalogue.yaml`, filtered for the section (S23 only on `/ng/`, S32 never on `/za/`). |
| `pages` | `section`*, `limit`, `cols` | Cards for the published pages in that section (this section's pages first, global fallback on English sections). |
| `sources` | `law`*, `ids` | A cited list of instruments from `data/law/<cc>.yaml`. |
| `law-table` | | The cross-country drone-law comparison (law hub). |
| `segments` | `data`* | The 500 m encroachment-density chart from `data/samples/<data>.json`. |
| `register` | `data`*, `limit` | The register excerpt table from the same file. |
| `details` | `summary`*, `open` (`true`) | A collapsible block. |
| `lead` | | Larger intro text. |
| `country-sites` | | Buttons to each country site that exists (renders nothing until one does). |

Icons (`icon="…"`): check, arrow-right, globe, pipeline, mine, power, clipboard, rail, sun, tree, building, corridor, boundary, shield, excavation, calendar, route, history, file-check, file-text, houses, compare, leaf, water, drone, satellite, target, map, layers, user-check, scale, search, mail, alert, info, external, clock, lock, download, ruler, flag, x-circle, send, grid, eye, language. Add new ones to `data/icons.yaml` (24×24, stroke style, no fills). Never use emoji.

Headings: write `## Heading {#anchor}` for a stable anchor (industry sections need stable anchors so solution pages can deep-link). Other headings get automatic ids. Every page has one H1 (from front matter); body headings start at `##` or come from a section `title`.

Tables: ordinary Markdown tables; the build wraps them in a scrollable box so they never widen the page on phones.

---

## 7. Links, calls to action and forms

- Internal links are **root-relative and clean**: `/methodology`, `/za/pipelines`, `/faq#drones`. Never `.html`, never relative (`../x`), never through a redirect. The build checks every link and every `#anchor`.
- `key:<cluster-key>` links resolve to that page in the current section, else the global page: `[resettlement baselines](key:resettlement-cut-off-baselines)`. The build fails if no page has the key.
- Contact links carry an intent: `/contact?intent=proposal` (default), `sample-report`, `tender`. Buttons with `intent:` and the closing band add `industry=`, `service=` and `country=` automatically so the form arrives pre-filled.
- The one CTA label site-wide is **Request a proposal** (PT: **Pedir proposta**).
- External links: to primary sources wherever you can (the gazette PDF, the regulator's page). Open data is credited, never hot-linked for imagery.

---

## 8. Images

1. Put the master (JPEG or PNG, as large as you have, up to 2400 px is used) in `site/images/…`, for example `site/images/samples/t9-km5-6.jpg`.
2. Use `:::figure{src="samples/t9-km5-6" alt="…" caption="…" credit="…"}`. The build makes AVIF, WebP and JPEG at 480–2400 px with `srcset`, width, height and lazy loading.
3. Alt text says what the image shows, for someone who cannot see it. Captions say what it is and where it comes from.

What may be shown:
- **The T-9 replacement pipeline and Pande (Inhambane) route and imagery**: yes, the owner has permission. Captions must be honest: manual marks are "Reviewed · manual marks", never "AI detection" or "live"; the operator's own plants and well pads are never "encroachment". Do not name the client company or use its logo.
- **Google, Bing or Esri basemap imagery**: only as an illustration, credited "Imagery © Google" (etc.), and never presented as delivered survey imagery. The T-9 views are built this way.
- **Drone orthophotos or mosaics of Mozambique**: not without the Lei n.º 6/2024 authorisation (art. 16(1)(c) makes unauthorised reproduction an infraction). Ask the owner first.
- No flags, coats of arms, regulator or client logos, stock photos of people, or images that identify anyone.

`site/tools/make_samples.py` rebuilds the T-9 images and `data/samples/t9.json` from the app's stored job (read-only). The old site's JPEGs had red and blue swapped; the tool draws from the raw GeoTIFFs, which are natural colour.

Social cards (1200×630) are drawn automatically for every page from `og.headline`, `og.subline` and the country; there is nothing to upload.

---

## 9. The guards (what fails the build)

The build checks the rendered text of every page: visible text, `<title>`, meta description, Open Graph text, image alt text and the JSON-LD strings. `ERROR` fails the build. A match is excused only when a negation word (not, never, no, without, não, nunca, sem…) appears earlier **in the same sentence**, so "Scheduled, never real-time" passes and "…without permission. We are licensed" does not.

| Guard | Fails on | Write instead |
|---|---|---|
| Pricing (owner rule 1) | Any currency amount (`US$ 1,200`, `R1 200 000`, `90 000 meticais`, `₦5,000`), "pricing", "price list", "prices from", "per-km rate", "preços" | Nothing about price. "Every proposal is scoped to your route or site." |
| Placeholders (rule 10) | `[OWNER`, `TBD`, `TODO`, `lorem`, `XXX`, leftover `{{ }}` or `:::` | Omit the fact. Never publish a stub. |
| Certificates (rule 2) | "we hold", "licensed", "certified", "ROC holder", "approved operator", "our UASOC/ROC/permit…", PT "somos licenciados", "detemos a licença"… | "Drone surveys are subject to the permits and authorisations each job requires." "Afridrone is working towards the operator approvals each country requires." Law pages state requirements ("an operator needs a UASOC"), never possession. |
| Overstatement (rules 3, 5, 9) | real-time, 24/7, around the clock, live detection/monitoring/imagery, instant, through cloud, continuous, surveillance, court-grade, forensic, admissible, PS5-compliant, replaces the census, every building/structure, guaranteed capture dates, track/count people or vehicles, intruders/invaders/illegal miners, thermal, LiDAR, multispectral, RTK, survey-grade, methane, leak detection, coming soon, on demand, roadmap, beta, client portal, login, dashboard, API, SLA, accuracy percentages, "results in 5 days"; PT tempo real, vigilância, contínuo, em breve, invasores, através das nuvens… | "Scheduled re-surveys", "after each survey", "flagged automatically and confirmed by a reviewer", "supports legal and community processes", "supports IFC Performance Standard 5 cut-off-date records", "structures, cleared ground, excavations, tracks". Radar may be described as showing larger changes "even under rainy-season cloud" (S14); optical imagery never sees through cloud. |
| pt-MZ vocabulary (rule 12) | Brazilian forms: monitoramento, equipe, contato, usuário, registro, conosco, celular, planejamento, gerenciamento, treinamento, aterrissagem, decolagem, controle de, de fato, loteamento, econômico/quilômetro-type accents, ótimo… | monitoria/monitorização, equipa, contacto, utilizador, registo, connosco, telemóvel, planeamento, gestão, formação, aterragem, descolagem, controlo de, de facto, económico, quilómetro, óptimo. |
| pt-MZ spelling (warn) | AO90 forms: proteção, projeto, setor, infraestrutura, ação, direção, eletricidade, detetar, atual, objetivo, ativo, afetado, seleção | The site uses pre-AO spelling (Boletim da República): protecção, projecto, sector, infra-estrutura, acção, direcção, electricidade, detectar, actual, objectivo, activo, afectado, selecção. |
| Structure | Missing or duplicate title/description, title > 65 or description outside 70–170 characters, not exactly one H1, missing canonical, missing og tags, `<img>` without width/height/alt, JSON-LD that does not parse or contains an Offer or price | Fix the front matter. |
| Links | Broken internal link, missing `#anchor`, `.html` links, relative links, `key:` that resolves nowhere | Use clean root-relative URLs or `key:` links. |
| Near-duplicates | Two same-language pages in different sections that share 70 % or more of their 5-word runs (warn at 50 %): cluster members, and country pages of the same template | Write country substance: local law, regulators, programmes, vocabulary, FAQs. A country page that cannot meet the minimum is not built. |
| Law pages | Missing not-legal-advice line, instruments without `id/title/identifier/date/url/last_checked`, non-https sources; warns when `last_reviewed` is over 120 days old | Keep `data/law/<cc>.yaml` complete and current. |

Do not weaken a rule in `rules.yaml` to get a page through: rewrite the page. If a rule is genuinely wrong, change it in its own commit with the reason, and run `--selftest`.

---

## 10. Country data and drone-law pages

**Rule 2 applies to every word.** Country drone-law pages state the requirements factually, **only from the verified briefs** (`regional/{mz,za,ng}.FINAL.md`), with a source link for each requirement, the "as of 26 September 2026" date (automatic on `law` pages) and the not-legal-advice line (automatic). Never say or imply that AfriScan or Afridrone holds any certificate, licence, permit or clearance. Anything the brief marks `[unverified]` stays off the site. Items the briefs say need counsel (for example whether satellite-only work falls outside Mozambique's Lei 6/2024) are stated as open, not answered.

`data/law/<cc>.yaml` schema (a worked example is `site/tests/fixtures/demo/law/za.yaml`):

```yaml
country_name: South Africa              # in English (PT pages show the same instrument titles)
last_reviewed: 2026-09-26               # when someone last re-checked every instrument
regulator: {name: South African Civil Aviation Authority (SACAA), url: https://www.caa.co.za/}
summary:                                # one row of the /drone-regulations comparison table
  regulator: South African Civil Aviation Authority (SACAA)
  instruments: Civil Aviation Regulations Part 101 (26th Amendment, 2023; 33rd Amendment, 2026)
instruments:                            # every instrument the page relies on
  - id: car-part101-26th                # used by :::sources{law="za" ids="car-part101-26th"}
    title: Civil Aviation Regulations, 26th Amendment (Part 101 substituted)
    identifier: GN R.3170, GG 48228     # gazette / notice / S.I. number
    date: 2023-03-17
    url: https://www.gov.za/sites/default/files/gcis_document/202303/48228rg11556gon3170.pdf
    last_checked: 2026-09-26
    note: Optional one-line note (inline Markdown).
```

Quote Mozambican laws in Boletim form: "Lei n.º 19/97, de 1 de Outubro (Lei de Terras)". On English Mozambique pages give the Portuguese title first with an English gloss.

Procurement and local-content pages list only registrations and facts the owner has confirmed. Where a fact is missing, say what a buyer can ask for in the proposal, or leave the line out.

---

## 11. Portuguese (pt-MZ)

- Write natively in Mozambican/European Portuguese as used in the Boletim da República, never Brazilian, and never raw machine translation. The glossary (`data/glossary/pt-MZ.yaml`) lists banned forms, AO90 spellings to avoid and the preferred vocabulary: zona de protecção parcial, faixa de servidão, servidão administrativa, linha de transporte de energia, gasoduto/oleoduto, construções (not invasões), machambas, benfeitorias, agregado familiar, DUAT, licença especial, reassentamento, PAR, data de corte, censo e inventário de bens, EIA, ortofotomapa, monitoria.
- Numbers use a decimal comma and a thin space for thousands (0,5 m; 14 500 km). Dates: 26 de Setembro de 2026.
- Prefer impersonal constructions or "a sua empresa"; the reviewer will settle the register.
- Interface strings are in `data/i18n/pt-MZ.yaml` (every key must exist in both files).
- Set `reviewed_on` and `reviewed_by_role` only after a real native review. The build warns on every PT page without them.
- The footer line "Falamos português" is not shown until someone can answer enquiries in Portuguese (owner question).

---

## 12. Facts you must not invent

Omit, never stub: the legal entity, registration numbers, NUIT, VAT, address, phone or WhatsApp, team names, insurance, data-storage location and retention, response times, supplier-portal registrations, memberships, partners, client names, case studies, logos, accuracy or turnaround figures. The contact form endpoint and email stay exactly as they are in `site.yaml`.

Brand: "AfriScan by Afridrone". Afridrone is AfriScan's sister drone-services brand (afridr.one). The Organization JSON-LD (in `site.yaml`) disambiguates AfriScan from Afriscan Construction (South Africa) and Afriscan Kenya.

Every service in `catalogue.yaml` is presented as available now. No "coming soon", "on demand", "future", "roadmap", "beta" or "pilot product" labels. Thermal, LiDAR, multispectral, RTK/survey-grade and gas-sensing services are not offered. Never offer a client portal, logins, dashboards or an API. A person reviews every result before delivery: say so.

---

## 13. Before you commit

- `site/build.py` shows **0 errors**; every warning is understood.
- `--selftest` passes if you touched `build.py`, `rules.yaml` or the glossary.
- Screenshots at 1440 and 390 read and fixed; `shoot.py` reports no overflow.
- `git diff --stat dist/` matches what you meant to change (the build is deterministic: rebuilding without source changes leaves `dist/` unchanged apart from sitemap dates of uncommitted files).
- Commit source and `dist/` together. Do not push or deploy: that is a separate step.
