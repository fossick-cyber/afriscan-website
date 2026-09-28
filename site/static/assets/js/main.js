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
      document.body.classList.toggle("region-open", d.open);
      if (!d.open) return;
      setMenu(false);
      closePanels();
      const current = $(".region-menu a[aria-current]", d);   // the menu scrolls: bring this site into view
      if (current) current.scrollIntoView({ block: "nearest" });
    });
    d.addEventListener("focusout", e => {
      if (e.relatedTarget && !d.contains(e.relatedTarget)) d.removeAttribute("open");
    });
  });
  // A mega-menu also opens on :hover (the same media query as in site.css), which pointerdown never
  // sees: close the country menu and any clicked-open panel first.
  const hoverPanels = window.matchMedia("(hover: hover) and (min-width: 1200px)");
  // A panel opened with its toggle closes when keyboard focus moves on past it (the drawer keeps its sections).
  const desktop = window.matchMedia("(min-width: 1200px)");
  $$(".nav-item.has-panel").forEach(li => {
    li.addEventListener("focusout", e => {
      if (desktop.matches && e.relatedTarget && !li.contains(e.relatedTarget)) closePanels();
    });
  });
  $$(".nav-item.has-panel").forEach(li => {
    li.addEventListener("pointerenter", e => {
      if (e.pointerType !== "mouse" || !hoverPanels.matches) return;
      closeRegion();
      closePanels(li);
    });
  });
  window.addEventListener("resize", () => { if (window.innerWidth >= 1200) setMenu(false); }, { passive: true });

  // ---------------------------------------------------------------- contact form (FormSubmit)
  const form = $("#contactForm");
  if (!form) return;
  // Links from industry, solution and country pages carry their context; it travels in hidden fields.
  const params = new URLSearchParams(location.search);
  const param = k => (params.get(k) || "").replace(/[^\w-]/g, "").slice(0, 80);
  for (const k of ["intent", "industry", "service"]) if (param(k)) form.elements[k].value = param(k);
  const country = form.elements.country;
  try {
    const name = JSON.parse(country.dataset.names || "{}")[param("country").toLowerCase()];
    if (name) country.value = name;
  } catch (_) { /* keep the page's own country */ }

  // Name and email are the only required fields; their messages come from the page's language.
  form.noValidate = true;
  const required = $$("[data-err-required]", form);
  const problem = el => !el.value.trim() ? el.dataset.errRequired
    : !el.validity.valid ? (el.dataset.errFormat || el.dataset.errRequired) : "";
  const flag = (el, text) => {
    const box = document.getElementById(el.getAttribute("aria-describedby"));
    if (text) el.setAttribute("aria-invalid", "true"); else el.removeAttribute("aria-invalid");
    if (box) { box.textContent = text; box.hidden = !text; }
  };
  required.forEach(el => el.addEventListener("input", () => { if (el.hasAttribute("aria-invalid")) flag(el, problem(el)); }));

  const msg = document.getElementById("formMsg");
  const button = $('button[type="submit"]', form);
  const label = button ? $("span", button) : null;
  const labelText = label ? label.textContent : "";

  form.addEventListener("submit", async e => {
    e.preventDefault();
    if (form.elements._honey.value) { location.assign(form.dataset.thanks); return; }
    const bad = required.filter(el => { const p = problem(el); flag(el, p); return p; });
    if (bad.length) { bad[0].focus(); return; }
    const data = {};
    for (const [k, v] of new FormData(form).entries()) {
      const s = String(v).trim();
      if (s || k.startsWith("_")) data[k] = s;
    }
    data._subject = `${form.dataset.subject}: ${data.name}${data.company ? " (" + data.company + ")" : ""}`;
    form.elements._subject.value = data._subject;
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
