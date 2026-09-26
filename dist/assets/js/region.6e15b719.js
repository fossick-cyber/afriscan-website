// Suggests the visitor's country site. Never redirects. Remembers the choice in this browser.
// Built by site/build.py: SITES lists only the country sections that exist.
(() => {
  const SITES = {"MZ": {"links": [["pt-MZ", "/mz/pt/", "Português"], ["en-MZ", "/mz/", "English"]]}, "ZA": {"links": [["en-ZA", "/za/", "South Africa"]]}, "NG": {"links": [["en-NG", "/ng/", "Nigeria"]]}};
  const TEXT = {
    MZ: { lang: "pt-MZ", text: "Está em Moçambique? Veja o site de Moçambique." },
    ZA: { lang: "en-ZA", text: "Looking for AfriScan South Africa?" },
    NG: { lang: "en-NG", text: "Looking for AfriScan Nigeria?" },
  };
  const store = (area, k, v) => {
    try { const s = window[area]; if (v === undefined) return s.getItem(k); s.setItem(k, v); } catch (e) { return null; }
    return null;
  };
  const CHOICE = "afriscan-region", GEO = "afriscan-geo";
  const here = document.documentElement.dataset.country || "";
  document.addEventListener("click", e => {
    const a = e.target.closest && e.target.closest("a[data-region]");
    if (a) store("localStorage", CHOICE, a.dataset.region);
  });
  if (store("localStorage", CHOICE)) return;
  const tzGuess = () => {
    try {
      return { "Africa/Maputo": "MZ", "Africa/Johannesburg": "ZA", "Africa/Lagos": "NG" }[Intl.DateTimeFormat().resolvedOptions().timeZone];
    } catch (e) { return undefined; }
  };
  const alt = code => {
    const l = document.querySelector(`link[rel="alternate"][hreflang="${code}" i]`);
    return l ? new URL(l.href).pathname : null;
  };
  const cached = store("sessionStorage", GEO);
  const geo = cached ? Promise.resolve(cached)
    : fetch("/geo", { cache: "no-store" }).then(r => (r.ok ? r.json() : {})).catch(() => ({}))
        .then(({ country }) => { const cc = SITES[country] ? country : (tzGuess() || "none");
                                 store("sessionStorage", GEO, cc); return cc; });
  geo.then(cc => {
    if (!SITES[cc] || !TEXT[cc] || cc === here) return;
    const bar = document.createElement("aside");
    bar.className = "region-banner";
    bar.setAttribute("aria-label", "Country site");
    const p = document.createElement("p");
    p.lang = TEXT[cc].lang; p.textContent = TEXT[cc].text;
    bar.append(p);
    for (const [code, home, label] of SITES[cc].links) {
      const a = document.createElement("a");
      a.href = alt(code) || home;
      a.hreflang = code; a.lang = code; a.textContent = label; a.className = "btn btn-sm";
      a.addEventListener("click", () => store("localStorage", CHOICE, cc));
      bar.append(a);
    }
    const stay = document.createElement("button");
    stay.type = "button"; stay.className = "region-banner-close"; stay.textContent = "×";
    stay.setAttribute("aria-label", "Stay on this site");
    stay.addEventListener("click", () => { store("localStorage", CHOICE, here || "GLOBAL"); bar.remove(); });
    bar.append(stay);
    document.body.append(bar);
  });
})();
