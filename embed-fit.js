/* Include at the end of any interactive figure that is shown inside a chapter (iframe):
   <script src="../../embed-fit.js"></script>
   It tells the chapter page how tall the figure is, so the frame fits it exactly
   (no blank space under the graph, no inner scrollbar). Does nothing outside an iframe. */
(function () {
  if (window.parent === window) return;
  var last = 0, full = false;
  function px(v) { return parseFloat(v) || 0; }
  function measure() {
    var bs = getComputedStyle(document.body), h = px(bs.paddingTop) + px(bs.paddingBottom);
    Array.prototype.forEach.call(document.body.children, function (el) {
      if (/^(SCRIPT|STYLE|LINK|TEMPLATE)$/.test(el.tagName)) return;
      var cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.position === 'fixed' || cs.position === 'absolute') return;
      h += el.getBoundingClientRect().height + px(cs.marginTop) + px(cs.marginBottom);
    });
    return Math.ceil(h) + 1;
  }
  function report() {
    if (full) return;
    var h = measure();
    if (h > 1 && h !== last) { last = h; window.parent.postMessage({ type: 'bioprax-embed-height', height: h }, '*'); }
  }
  window.addEventListener('message', function (e) {
    var d = e.data || {};
    if (d.type === 'bioprax-hud') { full = !!d.full; if (!full) report(); }
  });
  window.addEventListener('resize', report);
  window.addEventListener('load', report);
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(report);
  if (window.ResizeObserver) {
    var ro = new ResizeObserver(report);
    Array.prototype.forEach.call(document.body.children, function (el) { ro.observe(el); });
  }
  report();
})();
