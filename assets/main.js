// Mobile nav toggle
document.querySelector('.mobile-toggle').addEventListener('click', e => {
  const open = document.querySelector('.nav-links').classList.toggle('open');
  e.currentTarget.setAttribute('aria-expanded', open);
});

// Close mobile nav on link click
document.querySelectorAll('.nav-links a').forEach(link => {
  link.addEventListener('click', () => {
    document.querySelector('.nav-links').classList.remove('open');
    document.querySelector('.mobile-toggle').setAttribute('aria-expanded', 'false');
  });
});

// Contact form — delivered via FormSubmit (https://formsubmit.co)
document.getElementById('contactForm')?.addEventListener('submit', async e => {
  e.preventDefault();
  const form = e.target;
  const msg = document.getElementById('formMsg');
  const button = form.querySelector('button[type="submit"]');
  const data = Object.fromEntries(new FormData(form));
  data._subject = `New AfriScan enquiry — ${data.name || 'no name'}${data.company ? ' (' + data.company + ')' : ''}`;

  button.disabled = true;
  button.textContent = 'Sending…';
  msg.textContent = '';
  try {
    const res = await fetch('https://formsubmit.co/ajax/fossickpictures@gmail.com', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify(data)
    });
    const body = await res.json();
    if (!res.ok || String(body.success) !== 'true') throw new Error(body.message || `HTTP ${res.status}`);
    msg.textContent = 'Thanks! Your enquiry has been sent. We\'ll be in touch soon.';
    form.reset();
  } catch (err) {
    // fall back to a native POST so the enquiry still gets through
    console.error('AJAX submit failed, falling back to form POST:', err);
    form.submit();
    return;
  } finally {
    button.disabled = false;
    button.textContent = 'Send Enquiry';
  }
});

// Navbar background + hero scene parallax on scroll
const heroScene = document.querySelector('.hero-scene');
const reducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
window.addEventListener('scroll', () => {
  document.querySelector('.navbar').style.background =
    window.scrollY > 50
      ? 'rgba(13,17,23,.98)'
      : 'rgba(13,17,23,.92)';
  if (heroScene && !reducedMotion) {
    heroScene.style.transform = `translateY(${window.scrollY * 0.25}px)`;
  }
}, { passive: true });

// 3D tilt: the hero diorama leans toward the pointer
const stage = document.querySelector('.hero-scene .stage');
const hero = document.querySelector('.hero');
if (stage && hero && !reducedMotion && window.matchMedia('(pointer: fine)').matches) {
  let targetX = 0, targetY = 0, curX = 0, curY = 0, raf = null;
  const tick = () => {
    curX += (targetX - curX) * 0.06;
    curY += (targetY - curY) * 0.06;
    stage.style.transform = `rotateX(${curY.toFixed(3)}deg) rotateY(${curX.toFixed(3)}deg)`;
    if (Math.abs(targetX - curX) + Math.abs(targetY - curY) > 0.002) {
      raf = requestAnimationFrame(tick);
    } else {
      raf = null;
    }
  };
  const nudge = () => { if (!raf) raf = requestAnimationFrame(tick); };
  hero.addEventListener('mousemove', e => {
    const r = hero.getBoundingClientRect();
    targetX = ((e.clientX - r.left) / r.width - 0.5) * 5;
    targetY = ((e.clientY - r.top) / r.height - 0.5) * -3.5;
    nudge();
  });
  hero.addEventListener('mouseleave', () => { targetX = 0; targetY = 0; nudge(); });
}

// Secret: clear the red-flagged building. Desktop — type "doze".
// Touch/anywhere — triple-tap the red high-risk marker.
(() => {
  const scene = document.querySelector('.hero-scene');
  const bldg3 = document.getElementById('bldg3');
  if (!scene || !bldg3) return;
  let running = false;

  function demolish() {
    if (running) return;
    running = true;
    scene.classList.add('dozing');
    setTimeout(() => bldg3.classList.add('cleared'), 3900);   // label flips to CLEARED
    setTimeout(() => {                                         // reset, then huts creep back
      scene.classList.remove('dozing');
      bldg3.classList.remove('cleared');
      bldg3.classList.add('regrowing');
      setTimeout(() => { bldg3.classList.remove('regrowing'); running = false; }, 1500);
    }, 8500);
  }

  // typed keyword (desktop)
  const target = 'doze';
  let buf = '';
  document.addEventListener('keydown', e => {
    if (e.key.length !== 1 || e.metaKey || e.ctrlKey || e.altKey) return;
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select') return;
    buf = (buf + e.key.toLowerCase()).slice(-target.length);
    if (buf === target) demolish();
  });

  // triple-tap the red marker — location-based so overlapping layers can't block it
  let taps = 0, tapTimer = null;
  document.addEventListener('pointerup', e => {
    const r = bldg3.getBoundingClientRect();
    const pad = 12;
    const hit = e.clientX >= r.left - pad && e.clientX <= r.right + pad &&
                e.clientY >= r.top - pad && e.clientY <= r.bottom + pad;
    if (!hit) return;
    taps++;
    clearTimeout(tapTimer);
    tapTimer = setTimeout(() => { taps = 0; }, 700);
    if (taps >= 3) { taps = 0; demolish(); }
  }, { passive: true });
})();

// Scroll-reveal animations, staggered within each container
if ('IntersectionObserver' in window && !reducedMotion) {
  const targets = document.querySelectorAll(
    '.card, .feature, .step, .industry-card, .example-card, .highlight, .pricing-model-item'
  );
  const perParent = new Map();
  targets.forEach(el => {
    el.classList.add('reveal');
    const n = perParent.get(el.parentElement) || 0;
    el.style.transitionDelay = `${Math.min(n * 80, 480)}ms`;
    perParent.set(el.parentElement, n + 1);
  });
  const io = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('in-view');
        io.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });
  targets.forEach(el => io.observe(el));
}
