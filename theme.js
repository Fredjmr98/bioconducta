(function () {
  var r = document.documentElement;
  try { var t = localStorage.getItem('theme'); if (t) r.setAttribute('data-theme', t); } catch (e) {}

  var svg = function (inner) {
    return '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + inner + '</svg>';
  };
  var MOON = svg('<path d="M20.5 14.2A8.5 8.5 0 0 1 9.8 3.5a8.5 8.5 0 1 0 10.7 10.7z"/>');
  var SUN = svg('<circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M2.5 12h2M19.5 12h2M5.3 5.3l1.4 1.4M17.3 17.3l1.4 1.4M5.3 18.7l1.4-1.4M17.3 6.7l1.4-1.4"/>');

  document.addEventListener('DOMContentLoaded', function () {
    var b = document.getElementById('theme-toggle'); if (!b) return;
    function dark() { var t = r.getAttribute('data-theme'); return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches; }
    function sync() {
      var d = dark();
      b.innerHTML = d ? SUN : MOON;
      b.setAttribute('aria-pressed', d);
      b.setAttribute('aria-label', d ? 'Activar modo claro' : 'Activar modo oscuro');
      b.title = d ? 'Modo claro' : 'Modo oscuro';
      Array.prototype.forEach.call(document.querySelectorAll('iframe'), function (f) {
        try { f.contentWindow.postMessage({ type: 'bioprax-theme', dark: d }, '*'); } catch (e) {}
      });
    }
    b.addEventListener('click', function () {
      var n = dark() ? 'light' : 'dark'; r.setAttribute('data-theme', n);
      try { localStorage.setItem('theme', n); } catch (e) {}
      sync();
    });
    sync();
  });
})();
