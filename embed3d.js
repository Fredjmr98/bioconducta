/* Embedded interactive figure (<figure class="embed3d">): fullscreen button, theme sync with the frame.
   Native fullscreen when available; otherwise (e.g. iPhone) the frame fills the window. */
(function () {
  function dark() { var t = document.documentElement.getAttribute('data-theme'); return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches; }
  document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('.embed3d').forEach(function (fig) {
      var frame = fig.querySelector('.embed3d-frame'), iframe = frame.querySelector('iframe');
      var btn = fig.querySelector('[data-embed3d="fullscreen"]'), close = fig.querySelector('.embed3d-close');
      var rq = frame.requestFullscreen || frame.webkitRequestFullscreen, pseudo = false;
      function post(d) { try { iframe.contentWindow.postMessage(d, '*'); } catch (e) {} }
      function theme() { post({ type: 'bioprax-theme', dark: dark() }); }
      function isFull() { return pseudo || (document.fullscreenElement || document.webkitFullscreenElement) === frame; }
      function sync() { post({ type: 'bioprax-hud', full: isFull() }); theme(); }
      function enter() {
        if (rq) { rq.call(frame); }
        else { pseudo = true; fig.classList.add('is-full'); document.body.style.overflow = 'hidden'; sync(); }
      }
      function leave() {
        if (pseudo) { pseudo = false; fig.classList.remove('is-full'); document.body.style.overflow = ''; sync(); }
        else if (document.exitFullscreen) { document.exitFullscreen(); }
      }
      iframe.addEventListener('load', theme);
      btn.addEventListener('click', enter);
      close.addEventListener('click', leave);
      document.addEventListener('fullscreenchange', sync);
      document.addEventListener('webkitfullscreenchange', sync);
      document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && pseudo) leave(); });
    });
  });
  /* the frame's own theme button (visible in fullscreen) drives the page's button */
  window.addEventListener('message', function (e) {
    var d = e.data || {}, b = document.getElementById('theme-toggle');
    /* a figure that reports its own height (embed-fit.js) gets a frame that fits it exactly */
    if (d.type === 'bioprax-embed-height' && d.height > 0) {
      document.querySelectorAll('.embed3d-frame iframe').forEach(function (f) {
        if (f.contentWindow === e.source) { f.parentNode.setAttribute('data-auto', ''); f.parentNode.style.setProperty('--embed-h', Math.ceil(d.height) + 'px'); }
      });
      return;
    }
    if (d.type === 'bioprax-theme-changed' && b && dark() !== !!d.dark) b.click();
  });
})();
