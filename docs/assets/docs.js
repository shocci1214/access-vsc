// Access VSC ドキュメント 共通スクリプト（コピー、目次の現在位置、テーマ、学習の進み具合）
(function () {
  var root = document.documentElement;
  var store = {
    get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  };

  // テーマ
  var THEME = 'accessvsc-docs-theme';
  var saved = store.get(THEME);
  if (saved) root.setAttribute('data-theme', saved);
  document.querySelectorAll('[data-theme-toggle]').forEach(function (b) {
    b.addEventListener('click', function () {
      var cur = root.getAttribute('data-theme');
      var dark = cur ? cur === 'dark' : window.matchMedia('(prefers-color-scheme: dark)').matches;
      var next = dark ? 'light' : 'dark';
      root.setAttribute('data-theme', next);
      store.set(THEME, next);
    });
  });

  // コード のコピー
  document.querySelectorAll('pre').forEach(function (pre) {
    var b = document.createElement('button');
    b.type = 'button'; b.className = 'copy'; b.textContent = 'コピー';
    b.addEventListener('click', function () {
      var code = pre.querySelector('code');
      var text = code ? code.innerText : pre.innerText;
      if (navigator.clipboard) navigator.clipboard.writeText(text).then(function () {
        b.textContent = 'コピーしました';
        setTimeout(function () { b.textContent = 'コピー'; }, 1500);
      }, function () {});
    });
    pre.appendChild(b);
  });

  // 目次の現在位置
  var links = Array.prototype.slice.call(document.querySelectorAll('nav.toc a[href^="#"]'));
  var map = {};
  links.forEach(function (a) { map[a.getAttribute('href').slice(1)] = a; });
  if (links.length && 'IntersectionObserver' in window) {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (e.isIntersecting && map[e.target.id]) {
          links.forEach(function (a) { a.classList.remove('active'); });
          map[e.target.id].classList.add('active');
        }
      });
    }, { rootMargin: '-60px 0px -75% 0px' });
    Object.keys(map).forEach(function (id) { var el = document.getElementById(id); if (el) io.observe(el); });
  }

  // 到達目標のチェック（ブラウザに保存）と進み具合のバー
  var boxes = Array.prototype.slice.call(document.querySelectorAll('.objectives input[type=checkbox][data-key]'));
  function updateBars() {
    document.querySelectorAll('[data-progress]').forEach(function (bar) {
      var group = bar.getAttribute('data-progress');
      var items = boxes.filter(function (b) { return !group || b.getAttribute('data-key').indexOf(group) === 0; });
      var done = items.filter(function (b) { return b.checked; }).length;
      var pct = items.length ? Math.round(done * 100 / items.length) : 0;
      var fill = bar.querySelector('div');
      if (fill) fill.style.width = pct + '%';
      var label = document.querySelector('[data-progress-label="' + group + '"]');
      if (label) label.textContent = done + ' / ' + items.length;
    });
  }
  boxes.forEach(function (b) {
    var key = 'accessvsc-goal-' + b.getAttribute('data-key');
    b.checked = store.get(key) === '1';
    b.addEventListener('change', function () { store.set(key, b.checked ? '1' : '0'); updateBars(); });
  });
  updateBars();
})();
