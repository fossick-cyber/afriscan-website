// Suggests the visitor's country site. Never redirects. Remembers the choice in this browser.
// Built by site/build.py from data/locales.yaml: DATA lists only the live country sections (never a draft),
// with each country's banner text, its sites, and the time zones used when /geo gives no country.
(() => {
  const DATA = {"sites": {"KE": {"lang": "en-KE", "links": [["en-KE", "/ke/", "Kenya"]], "text": "In Kenya? See our Kenya site."}, "MZ": {"lang": "pt-MZ", "links": [["pt-MZ", "/mz/pt/", "Português"], ["en-MZ", "/mz/", "English"]], "text": "Está em Moçambique? Veja o site de Moçambique."}, "NG": {"lang": "en-NG", "links": [["en-NG", "/ng/", "Nigeria"]], "text": "In Nigeria? See our Nigeria site."}, "TZ": {"lang": "en-TZ", "links": [["en-TZ", "/tz/", "Tanzania"]], "text": "In Tanzania? See our Tanzania site."}, "UG": {"lang": "en-UG", "links": [["en-UG", "/ug/", "Uganda"]], "text": "In Uganda? See our Uganda site."}, "ZA": {"lang": "en-ZA", "links": [["en-ZA", "/za/", "South Africa"]], "text": "In South Africa? See our South Africa site."}, "ZM": {"lang": "en-ZM", "links": [["en-ZM", "/zm/", "Zambia"]], "text": "In Zambia? See our Zambia site."}}, "tz": {"Africa/Dar_es_Salaam": "TZ", "Africa/Johannesburg": "ZA", "Africa/Kampala": "UG", "Africa/Lagos": "NG", "Africa/Lusaka": "ZM", "Africa/Maputo": "MZ", "Africa/Nairobi": "KE"}, "ui": {"en": {"bar": "Country site", "stay": "Stay on this site"}, "en-GB": {"bar": "Country site", "stay": "Stay on this site"}, "en-KE": {"bar": "Country site", "stay": "Stay on this site"}, "en-MZ": {"bar": "Country site", "stay": "Stay on this site"}, "en-NG": {"bar": "Country site", "stay": "Stay on this site"}, "en-TZ": {"bar": "Country site", "stay": "Stay on this site"}, "en-UG": {"bar": "Country site", "stay": "Stay on this site"}, "en-ZA": {"bar": "Country site", "stay": "Stay on this site"}, "en-ZM": {"bar": "Country site", "stay": "Stay on this site"}, "pt-MZ": {"bar": "Site do país", "stay": "Ficar neste site"}}};
  const SITES = DATA.sites, TZ = DATA.tz;
  const L = document.documentElement.lang || "";
  const UI = DATA.ui[L] || DATA.ui[Object.keys(DATA.ui).find(k => k.slice(0, 2) === L.slice(0, 2))] || DATA.ui.en;
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
    try { return TZ[Intl.DateTimeFormat().resolvedOptions().timeZone]; } catch (e) { return undefined; }
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
    if (!SITES[cc] || cc === here) return;
    const site = SITES[cc];
    const bar = document.createElement("aside");
    bar.className = "region-banner";
    bar.setAttribute("aria-label", UI.bar);
    const p = document.createElement("p");
    p.lang = site.lang; p.textContent = site.text;
    bar.append(p);
    for (const [code, home, label] of site.links) {
      const a = document.createElement("a");
      a.href = alt(code) || home;
      a.hreflang = code; a.lang = code; a.textContent = label; a.className = "btn btn-sm";
      a.addEventListener("click", () => store("localStorage", CHOICE, cc));
      bar.append(a);
    }
    const stay = document.createElement("button");
    stay.type = "button"; stay.className = "region-banner-close"; stay.textContent = "×";
    stay.setAttribute("aria-label", UI.stay);
    stay.addEventListener("click", () => { store("localStorage", CHOICE, here || "GLOBAL"); bar.remove(); });
    bar.append(stay);
    document.body.append(bar);
  });
})();
