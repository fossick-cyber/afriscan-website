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
| `functions/geo.js` | `GET /geo` returns the visitor's country for the country-site banner |
| `functions/_middleware.js` | Host redirects: `www.afri-scan.com` and `afriscan-website.pages.dev` → `https://afri-scan.com` (301, path and query kept); preview hosts pass through. `dist/_routes.json` keeps `/assets/*` out of Functions |
| `wrangler.toml` | Pages project settings read on every Git build: `pages_build_output_dir = "./dist"` |
| `dist/` | Generated output. Committed. Never edited by hand |

## Deploying (not done from this repo's working copy without the owner's go-ahead)

- Cloudflare Pages project `afriscan-website`, production branch `main`.
- **`wrangler.toml` sets the build output directory to `dist`.** Cloudflare reads it on every Git build (build system v2 or later; the project is on v3) and it then overrides the dashboard setting, which on 2026-09-26 was still the repo root. With no build command, Pages serves the committed `dist/` and deploys `functions/` from the repo root. `site/build.py` fails if `wrangler.toml` goes missing or stops pointing at `./dist`. After the first push, check that `/site/build.py` returns 404 (architecture.FINAL §7.6).
- Host redirects (www → apex, `afriscan-website.pages.dev` → apex) are done by `functions/_middleware.js`, because the deploy token cannot create Bulk Redirects or Redirect Rules. If a Bulk Redirect list is ever set up in the dashboard (architecture.FINAL §7.2), delete the middleware and the `_routes.json` step in `site/build.py`. `_redirects` and `_headers` still apply to pages the middleware passes through (checked with `wrangler pages dev` and live on 2026-09-26). Path redirects are generated into `dist/_redirects`; security and cache headers into `dist/_headers`.
- IndexNow: the build publishes `/<indexnow_key>.txt` (key in `site/data/site.yaml`). After a deploy, ping `https://api.indexnow.org/indexnow` with the changed URLs.
- Before pushing: `site/build.py` must report 0 errors, and `git diff --exit-code dist/` after a rebuild must be clean.

`/assets/*` is served with `Cache-Control: immutable` for a year. CSS, JS, images and social cards get content-hashed names automatically; the font (`assets/fonts/`) and logo (`assets/brand/`) files do not, so give them a new file name if they ever change.
