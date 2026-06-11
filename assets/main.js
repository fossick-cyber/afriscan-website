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
