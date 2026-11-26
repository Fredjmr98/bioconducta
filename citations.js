/* Smart author-year citations. The year links to the reference entry, which links back to each place it was cited.
   In the text:      <span class="cite" data-ref="hoffmeyer2008"></span>                  -> Hoffmeyer 2008
                     <span class="cite" data-ref="hoffmeyer2008" data-only="year"></span>  -> 2008  (write "Hoffmeyer (" ... ")" by hand)
                     <span class="cite" data-ref="a2008,b2021" data-loc="p. 12"></span>   -> A 2008; B 2021, p. 12
   In the reference list:
                     <div class="ref" id="ref-hoffmeyer2008" data-author="Hoffmeyer" data-year="2008">Full reference…</div> */
document.addEventListener('DOMContentLoaded', function () {
  var refs = {}, uses = {};
  document.querySelectorAll('.ref[id^="ref-"]').forEach(function (r) { refs[r.id.slice(4)] = r; });

  document.querySelectorAll('.cite[data-ref]').forEach(function (c) {
    var keys = c.getAttribute('data-ref').split(/\s*,\s*/);
    var onlyYear = c.getAttribute('data-only') === 'year';
    var loc = c.getAttribute('data-loc');
    c.textContent = '';
    keys.forEach(function (key, i) {
      var r = refs[key];
      if (i) c.appendChild(document.createTextNode('; '));
      if (!r) {
        var m = document.createElement('span');
        m.className = 'cite-missing';
        m.textContent = '[' + key + '?]';
        c.appendChild(m);
        return;
      }
      if (!onlyYear) c.appendChild(document.createTextNode(r.getAttribute('data-author') + ' '));
      var n = uses[key] = (uses[key] || 0) + 1;
      var a = document.createElement('a');
      a.id = 'cite-' + key + '-' + n;
      a.href = '#ref-' + key;
      a.title = r.textContent.replace(/\s+/g, ' ').trim();
      a.textContent = r.getAttribute('data-year');
      c.appendChild(a);
    });
    if (loc) c.appendChild(document.createTextNode(', ' + loc));
  });

  Object.keys(uses).forEach(function (key) {
    var span = document.createElement('span');
    span.className = 'backrefs';
    for (var n = 1; n <= uses[key]; n++) {
      var b = document.createElement('a');
      b.href = '#cite-' + key + '-' + n;
      b.textContent = '\u21A9';
      b.setAttribute('aria-label', 'Volver al texto');
      span.appendChild(document.createTextNode(' '));
      span.appendChild(b);
    }
    refs[key].appendChild(span);
  });

  // a link opened with a #cite-... hash exists only after the script ran
  if (location.hash.indexOf('#cite-') === 0) {
    var t = document.getElementById(location.hash.slice(1));
    if (t) t.scrollIntoView({ block: 'center' });
  }
});
