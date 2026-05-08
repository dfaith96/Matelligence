(() => {
  const year = document.getElementById('year');
  if (year) year.textContent = String(new Date().getFullYear());

  const navToggle = document.getElementById('navToggle');
  const drawer = document.getElementById('navDrawer');
  const setOpen = (open) => {
    if (!navToggle || !drawer) return;
    navToggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    drawer.hidden = !open;
  };

  if (navToggle && drawer) {
    navToggle.addEventListener('click', () => {
      const open = navToggle.getAttribute('aria-expanded') === 'true';
      setOpen(!open);
    });

    drawer.addEventListener('click', (e) => {
      const target = e.target;
      if (target && target.tagName === 'A') setOpen(false);
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') setOpen(false);
    });

    document.addEventListener('click', (e) => {
      const t = e.target;
      if (!t) return;
      if (drawer.hidden) return;
      if (drawer.contains(t) || navToggle.contains(t)) return;
      setOpen(false);
    });
  }

  const form = document.getElementById('leadForm');
  if (form) {
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const data = new FormData(form);
      const name = String(data.get('name') || '').trim();
      const email = String(data.get('email') || '').trim();
      const company = String(data.get('company') || '').trim();
      const message = String(data.get('message') || '').trim();

      const subject = encodeURIComponent('Matelligence demo request');
      const body = encodeURIComponent(
        `Name: ${name}\nEmail: ${email}\nCompany: ${company}\n\nUse case:\n${message}`
      );

      window.location.href = `mailto:hello@matelligence.example?subject=${subject}&body=${body}`;
    });
  }
})();
