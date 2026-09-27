// AfriScan site script: navigation, country selector and the contact form. No tracking.
(() => {
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));

  // ---------------------------------------------------------------- mobile menu
  const toggle = $(".menu-toggle");
  const closeRegion = () => $$("details.region[open]").forEach(d => d.removeAttribute("open"));
  const setMenu = open => {
    if (!toggle) return;
    if (open) closeRegion();
    document.body.classList.toggle("menu-open", open);
    toggle.setAttribute("aria-expanded", String(open));
    toggle.setAttribute("aria-label", open ? toggle.dataset.labelClose : toggle.dataset.labelOpen);
  };
  if (toggle) toggle.addEventListener("click", () => setMenu(toggle.getAttribute("aria-expanded") !== "true"));

  // ---------------------------------------------------------------- dropdown panels
  const closePanels = except => {
    $$(".nav-item.open").forEach(li => {
      if (li === except) return;
      li.classList.remove("open");
      const b = $(".nav-toggle", li);
      if (b) b.setAttribute("aria-expanded", "false");
    });
  };
  $$(".nav-toggle").forEach(btn => {
    btn.addEventListener("click", () => {
      const li = btn.closest(".nav-item");
      const open = !li.classList.contains("open");
      closePanels(li);
      li.classList.toggle("open", open);
      btn.setAttribute("aria-expanded", String(open));
    });
  });

  document.addEventListener("keydown", e => {
    if (e.key !== "Escape") return;
    const openItem = $(".nav-item.open");
    const region = $("details.region[open]");
    if (openItem) {
      closePanels();
      const b = $(".nav-toggle", openItem);
      if (b) b.focus();
    } else if (region) {
      region.removeAttribute("open");
      $("summary", region).focus();
    } else if (document.body.classList.contains("menu-open")) {
      setMenu(false);
      toggle.focus();
    }
  });
  // pointerdown, not click: iOS Safari sends no click to the document for a tap on plain content.
  document.addEventListener("pointerdown", e => {
    if (!e.target.closest(".nav-item")) closePanels();
    if (!e.target.closest("details.region")) closeRegion();
    if (document.body.classList.contains("menu-open") && !e.target.closest(".site-header")) setMenu(false);
  });
  $$("details.region").forEach(d => {
    d.addEventListener("toggle", () => {
      if (!d.open) return;
      setMenu(false);
      closePanels();
    });
    d.addEventListener("focusout", e => {
      if (e.relatedTarget && !d.contains(e.relatedTarget)) d.removeAttribute("open");
    });
  });
  window.addEventListener("resize", () => { if (window.innerWidth >= 1200) setMenu(false); }, { passive: true });

  // ---------------------------------------------------------------- contact form (FormSubmit)
  const form = $("#contactForm");
  if (!form) return;
  const params = new URLSearchParams(location.search);
  const setVal = (id, v) => { const el = document.getElementById(id); if (el && v) el.value = v.slice(0, 80); };
  setVal("f-intent", params.get("intent"));
  setVal("f-industry", params.get("industry"));
  setVal("f-service", params.get("service"));
  const cc = (params.get("country") || "").toLowerCase();
  const country = document.getElementById("f-country");
  if (country && cc) {
    const opt = $$("option", country).find(o => o.dataset.code === cc);
    if (opt) opt.selected = true;
  }
  const assetIndex = { "oil-gas": 1, "power-utilities": 2, "rail-roads": 3, "mining": 4 }[params.get("industry")];
  const asset = document.getElementById("f-asset");
  if (asset && assetIndex && asset.options[assetIndex]) asset.selectedIndex = assetIndex;

  const msg = document.getElementById("formMsg");
  const button = $('button[type="submit"]', form);
  const label = button ? $("span", button) : null;
  const labelText = label ? label.textContent : "";

  form.addEventListener("submit", async e => {
    e.preventDefault();
    if ($('input[name="_honey"]', form).value) { location.assign(form.dataset.thanks); return; }
    const fd = new FormData(form);
    const data = {};
    for (const [k, v] of fd.entries()) data[k] = data[k] ? `${data[k]}, ${v}` : v;
    data._subject = `${form.dataset.subject}: ${data.name || ""}${data.organisation ? " (" + data.organisation + ")" : ""}`;
    if (button) button.disabled = true;
    if (label) label.textContent = form.dataset.sending;
    msg.textContent = "";
    msg.className = "form-status";
    try {
      const res = await fetch(form.dataset.ajax, {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify(data),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok || String(body.success) !== "true") throw new Error(body.message || `HTTP ${res.status}`);
      msg.textContent = form.dataset.success;
      msg.classList.add("is-ok");
      location.assign(form.dataset.thanks);
    } catch (err) {
      // Fall back to a normal POST: FormSubmit then redirects to our thank-you page (_next).
      msg.textContent = form.dataset.error;
      msg.classList.add("is-error");
      setTimeout(() => form.submit(), 600);
    } finally {
      if (button) button.disabled = false;
      if (label) label.textContent = labelText;
    }
  });
})();
