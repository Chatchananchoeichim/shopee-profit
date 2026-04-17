// Global Event Listeners & Initializations
document.addEventListener('DOMContentLoaded', () => {
  // Initialize Theme
  if (state.theme === 'dark') {
    document.documentElement.setAttribute('data-theme', 'dark');
    const icon = document.querySelector('.theme-toggle .material-symbols-rounded');
    if (icon) icon.innerText = 'light_mode';
  }

  // Initialize Auto-Animate
  if (typeof autoAnimate !== 'undefined') {
    const containers = [
      document.getElementById('missing-variants')
    ];
    containers.forEach(el => { if(el) autoAnimate(el); });
  }

  // Drag & Drop Handlers
  ['drop-order','drop-income','drop-ads','drop-stats'].forEach(id=>{
    const el = document.getElementById(id);
    if(!el) return;
    el.addEventListener('dragover', e=>{e.preventDefault(); el.style.borderColor='var(--primary)';});
    el.addEventListener('dragleave', ()=>el.style.borderColor='');
    el.addEventListener('drop', e=>{
      e.preventDefault(); el.style.borderColor='';
      const file = e.dataTransfer.files[0];
      if(!file) return;
      if(id === 'drop-order') handleFile({files:[file]}, 'order');
      if(id === 'drop-income') handleFile({files:[file]}, 'income');
      if(id === 'drop-ads') handleAdsFile({files:[file]});
      if(id === 'drop-stats') handleStatsFile({files:[file]});
    });
  });

  // Global Tippy Delegate for dynamic tooltips
  if (typeof tippy !== 'undefined') {
    tippy('body', {
      target: '[title]',
      animation: 'shift-away',
      arrow: true,
      theme: 'light-border',
      allowHTML: true,
      content(reference) {
        const title = reference.getAttribute('title');
        reference.removeAttribute('title');
        return title;
      },
    });
  }

  // Set User Initials
  const email = localStorage.getItem('shopee_user_email') || 'Guest';
  document.getElementById('user-email-sidebar').innerText = email;
  document.getElementById('user-initials').innerText = email.charAt(0).toUpperCase();

  // Initial Tab
  switchTab('import');
});
