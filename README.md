# afri-scan.com

The website for **AfriScan by Afridrone**: remote land and corridor monitoring for pipelines, power lines, concessions and project sites in Mozambique, South Africa and Nigeria.

The site is generated. Sources live in `site/`; the build writes the complete static site to `dist/`, which is what Cloudflare Pages serves.

```bash
/opt/favhousecheck/.venv/bin/python3 site/build.py              # build dist/ and run the house-rule guards
/opt/favhousecheck/.venv/bin/python3 site/build.py --selftest   # check that every guard still fails on seeded mistakes
python3 -m http.server 5091 --directory dist --bind 127.0.0.1   # preview (open /results.html etc. locally)
```

Writers: read **[CONTENT_GUIDE.md](CONTENT_GUIDE.md)** before adding or editing pages.

## Layout

| Path | What it is |
|---|---|
| `site/build.py`, `site/lib/` | Generator: Markdown + YAML front matter → Jinja2 templates → `dist/`; hreflang, JSON-LD, sitemaps, redirects, headers, responsive images, social cards, favicons; the guards |
| `site/content/` | One file per page, per section: `global/`, `mz/en/`, `mz/pt/`, `za/`, `ng/` |
| `site/data/` | Facts and strings: site/organisation, locales, catalogue, drone-law instruments, i18n, pt-MZ glossary, guard rules, redirects, icons, sample data |
| `site/templates/`, `site/static/` | Page templates, components, CSS, JS, self-hosted Inter font (OFL) |
| `site/images/` | Image masters (the T-9 sample views are rebuilt from the source GeoTIFFs by `site/tools/make_samples.py`) |
| `functions/geo.js` | The only Pages Function: `GET /geo` returns the visitor's country for the country-site banner |
| `dist/` | Generated output. Committed. Never edited by hand |

## Deploying (not done from this repo's working copy without the owner's go-ahead)

- Cloudflare Pages project `afriscan-website`, production branch `main`.
- **Build output directory must be `dist`** (the project currently serves the repo root). With no build command, Pages serves the committed `dist/` and deploys `functions/` from the repo root.
- Host redirects (www → apex, `afriscan-website.pages.dev` → apex) belong in Cloudflare Bulk Redirects, not in this repo. Path redirects are generated into `dist/_redirects`; security and cache headers into `dist/_headers`.
- Before pushing: `site/build.py` must report 0 errors, and `git diff --exit-code dist/` after a rebuild must be clean.

`/assets/*` is served with `Cache-Control: immutable` for a year. CSS, JS, images and social cards get content-hashed names automatically; the font (`assets/fonts/`) and logo (`assets/brand/`) files do not, so give them a new file name if they ever change.
