/* LaTeX via MathJax 3. Include once per page that has math: <script src="math.js"></script>
   Inline: $...$ or \( ... \)   Display: $$...$$ or \[ ... \]
   Numbered: \begin{equation}\label{eq:name} ... \end{equation}  (cite with \eqref{eq:name})
   Numbers show as (1.1), (1.2)... on the right; the prefix comes from <body data-chapter="1">. */
window.MathJax = {
  tex: {
    inlineMath: [['$', '$'], ['\\(', '\\)']],
    displayMath: [['$$', '$$'], ['\\[', '\\]']],
    tags: 'ams',
    tagSide: 'right',
    tagIndent: '0em',
    tagformat: {
      number: function (n) {
        var c = document.body && document.body.getAttribute('data-chapter');
        return (c ? c + '.' : '') + n;
      }
    }
  },
  options: { skipHtmlTags: ['script', 'noscript', 'style', 'textarea', 'pre', 'code'] }
};
(function () {
  var s = document.createElement('script');
  s.async = true;
  s.src = 'https://cdn.jsdelivr.net/npm/mathjax@3.2.2/es5/tex-chtml.js';
  document.head.appendChild(s);
})();
