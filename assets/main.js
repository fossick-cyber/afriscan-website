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
document.getElementById('contactForm').addEventListener('submit', async e => {
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
