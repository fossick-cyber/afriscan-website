#!/usr/bin/env python3
"""Screenshot built pages at desktop and phone widths, and report layout problems.

    /opt/favhousecheck/.venv/bin/python3 site/tools/shoot.py http://127.0.0.1:5091 OUT_DIR /results /features ...

The local preview (python3 -m http.server) does not map /x to x.html the way Cloudflare Pages
does, so paths are requested as /x.html. Prints horizontal overflow and images without size.
"""
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright

WIDTHS = {"1440": (1440, 900), "390": (390, 844)}


def local(path):
    if path.endswith("/"):
        return path + "index.html"
    return path + ".html"


def main():
    base, out = sys.argv[1].rstrip("/"), Path(sys.argv[2])
    paths = sys.argv[3:] or ["/"]
    out.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        for name, (w, h) in WIDTHS.items():
            ctx = browser.new_context(viewport={"width": w, "height": h}, device_scale_factor=1)
            page = ctx.new_page()
            for p in paths:
                page.goto(base + local(p), wait_until="networkidle")
                page.evaluate("""async () => { for (let y = 0; y < document.body.scrollHeight; y += 600) {
                    window.scrollTo(0, y); await new Promise(r => setTimeout(r, 150)); } window.scrollTo(0, 0); }""")
                page.wait_for_load_state("networkidle")
                page.wait_for_timeout(900)
                stem = (p.strip("/").replace("/", "_") or "home") + f"-{name}"
                page.screenshot(path=str(out / f"{stem}-fold.png"))
                page.screenshot(path=str(out / f"{stem}-full.png"), full_page=True)
                ov = page.evaluate("""() => {
                    const W = document.documentElement.clientWidth, bad = [];
                    for (const el of document.querySelectorAll('body *')) {
                        const r = el.getBoundingClientRect();
                        if (r.right > W + 1 && getComputedStyle(el).position !== 'fixed' && !el.closest('.table-wrap, .segments-chart'))
                            bad.push(el.tagName.toLowerCase() + '.' + (el.className || '').toString().split(' ')[0] + ' ' + Math.round(r.right));
                    }
                    return {scrollW: document.documentElement.scrollWidth, W, bad: bad.slice(0, 8)};
                }""")
                if ov["scrollW"] > ov["W"] or ov["bad"]:
                    print(f"OVERFLOW {p} @{name}: scrollWidth {ov['scrollW']} > {ov['W']}: {ov['bad']}")
                print(f"shot {stem}")
            ctx.close()
        browser.close()


if __name__ == "__main__":
    main()
