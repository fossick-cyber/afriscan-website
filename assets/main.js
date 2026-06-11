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

// Contact form — local storage for now (no backend)
document.getElementById('contactForm').addEventListener('submit', e => {
  e.preventDefault();
  const data = Object.fromEntries(new FormData(e.target));
  console.log('Enquiry submitted:', data);

  // Store locally
  const enquiries = JSON.parse(localStorage.getItem('afriscan_enquiries') || '[]');
  enquiries.push({ ...data, timestamp: new Date().toISOString() });
  localStorage.setItem('afriscan_enquiries', JSON.stringify(enquiries));

  document.getElementById('formMsg').textContent =
    'Thanks! Your enquiry has been saved. We\'ll be in touch soon.';
  e.target.reset();
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
