// Mobile nav toggle
document.querySelector('.mobile-toggle').addEventListener('click', () => {
  document.querySelector('.nav-links').classList.toggle('open');
});

// Close mobile nav on link click
document.querySelectorAll('.nav-links a').forEach(link => {
  link.addEventListener('click', () => {
    document.querySelector('.nav-links').classList.remove('open');
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

// Navbar background on scroll
window.addEventListener('scroll', () => {
  document.querySelector('.navbar').style.background =
    window.scrollY > 50
      ? 'rgba(13,17,23,.98)'
      : 'rgba(13,17,23,.92)';
});
