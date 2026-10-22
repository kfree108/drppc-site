/* Free Dr. PPC tools (drppc.ai/free/). Everything runs in the browser; nothing is sent anywhere. */
(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
  var num = function (id) { var el = document.getElementById(id); var v = parseFloat(String(el && el.value).replace(/[$,%\s]/g, '')); return isFinite(v) ? v : 0; };
  var money = function (v) { return (v < 0 ? '−$' : '$') + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  var money0 = function (v) { return (v < 0 ? '−$' : '$') + Math.round(Math.abs(v)).toLocaleString('en-US'); };
  var pct = function (v, d) { return (v * 100).toFixed(d == null ? 1 : d).replace(/\.0$/, '') + '%'; };
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  function csv(name, rows) {
    var text = rows.map(function (r) { return r.map(function (c) { c = String(c == null ? '' : c); return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c; }).join(','); }).join('\n');
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type: 'text/csv' }));
    a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
  }
  function live(ids, fn) { ids.forEach(function (id) { var el = document.getElementById(id); if (el) { el.addEventListener('input', fn); el.addEventListener('change', fn); } }); fn(); }

  /* ---------- 1. The 3× Rule ---------- */
  // Negate line = 3 × clicks per order. P(a term that truly converts at c shows 0 orders in n clicks) = (1 − c)^n.
  function line(p) { return Math.ceil(300 / p - 1e-9); }
  function rule() {
    var wrap = $('#t-rule'); if (!wrap) return;
    $('#r-calc').addEventListener('click', function () {
      var c = num('r-clicks'), o = num('r-orders');
      if (c > 0) { $('#r-cr').value = (o / c * 100).toFixed(1).replace(/\.0$/, ''); $('#r-cr').dispatchEvent(new Event('input')); }
    });
    function draw() {
      var p = Math.min(Math.max(num('r-cr'), 0.5), 60), n = line(p), miss = Math.pow(1 - p / 100, n);
      var term = num('r-term');
      $('#r-out').innerHTML =
        '<div><b>' + (100 / p).toFixed(1).replace(/\.0$/, '') + '</b><span>clicks per order, on average</span></div>' +
        '<div><b>' + n + ' clicks</b><span>your line: past this with no sale, it’s out</span></div>' +
        '<div><b>' + pct(1 - miss, 0) + ' sure</b><span>a term this good would show a sale by then ' + pct(1 - miss, 1) + ' of the time</span></div>' +
        (term > 0 ? '<div class="' + (term >= n ? 'warn' : '') + '"><b>' + (term >= n ? 'Negative exact' : 'Let it run') + '</b><span>' + term + ' clicks, no orders: ' + (term >= n ? 'past your line' : (n - term) + ' clicks to go before you judge it') + '</span></div>' : '');
      $('#r-math').innerHTML = '1 ÷ <em>' + p + '%</em> = ' + (100 / p).toFixed(2) + ' clicks per order · × 3 = <em>' + n + ' clicks</em> · chance a ' + p + '% term still shows 0 orders: (1 − ' + (p / 100) + ')<sup>' + n + '</sup> = ' + pct(miss, 1);
      var hit = Math.round(p);
      $$('#r-table tbody tr').forEach(function (tr) { tr.classList.toggle('hit', +tr.dataset.p === hit); });
    }
    var rows = [];
    for (var p = 1; p <= 40; p++) {
      var n = line(p), miss = Math.pow(1 - p / 100, n);
      rows.push('<tr data-p="' + p + '"><td class="n">' + p + '%</td><td class="n">' + (100 / p).toFixed(1).replace(/\.0$/, '') + '</td><td class="n"><b>' + n + '</b></td><td class="n">' + pct(miss, 1) + '</td><td class="n">' + pct(1 - miss, 1) + '</td></tr>');
    }
    $('#r-table tbody').innerHTML = rows.join('');
    $('#r-csv').addEventListener('click', function () {
      var out = [['Conversion rate', 'Clicks per order', 'Negate after (clicks, 0 orders)', 'Chance a good term still shows 0 orders', 'How sure']];
      for (var p = 1; p <= 40; p++) { var n = line(p), m = Math.pow(1 - p / 100, n); out.push([p + '%', (100 / p).toFixed(1), n, pct(m, 1), pct(1 - m, 1)]); }
      csv('dr-ppc-3x-rule-table.csv', out);
    });
    live(['r-cr', 'r-term'], draw);
  }

  /* ---------- 2. The 3-Number Bid ---------- */
  // Break-even ACoS = (price − costs before ads) ÷ price. Max bid = price × target ACoS × conversion rate.
  function bid() {
    if (!$('#t-bid')) return;
    function draw() {
      var price = num('b-price'), ref = price * num('b-ref') / 100;
      var costs = num('b-cogs') + num('b-freight') + ref + num('b-fba') + num('b-storage') + num('b-returns');
      var margin = price - costs, be = price > 0 ? margin / price : 0;
      var target = num('b-target') / 100, cr = num('b-cr') / 100, boost = num('b-boost') / 100;
      var max = price * target * cr, base = max / (1 + boost);
      var over = target > be;
      $('#b-out').innerHTML =
        '<div><b>' + pct(be, 1) + '</b><span>break-even ACoS: ' + money(margin) + ' left of ' + money(price) + ' before ads</span></div>' +
        '<div class="' + (over ? 'warn' : '') + '"><b>' + pct(target, 1) + '</b><span>' + (over ? 'your target is above break-even: every ad sale loses money. Only on a planned launch.' : (Math.abs(target - be) < 0.0005 ? 'right at break-even: a launch setting' : 'under break-even, so each ad sale makes money')) + '</span></div>' +
        '<div><b>' + money(max) + '</b><span>the most to pay for one click (your ceiling)</span></div>' +
        '<div><b>' + money(base) + '</b><span>base bid to type in, with a ' + pct(boost, 0) + ' top-of-search boost on top</span></div>';
      $('#b-math').innerHTML = 'Referral ' + money(ref) + ' · costs before ads ' + money(costs) + ' · break-even ' + money(margin) + ' ÷ ' + money(price) + ' = <em>' + pct(be, 1) + '</em><br>' +
        money(price) + ' × ' + pct(target, 1) + ' × ' + pct(cr, 1) + ' = <em>' + money(max) + '</em> ceiling · ÷ ' + (1 + boost).toFixed(2) + ' = <em>' + money(base) + '</em> base bid';
      var rows = [];
      for (var p = 1; p <= 25; p++) {
        var m = price * target * p / 100;
        rows.push('<tr' + (Math.round(cr * 100) === p ? ' class="hit"' : '') + '><td class="n">' + p + '%</td><td class="n"><b>' + money(m) + '</b></td><td class="n">' + money(m / (1 + boost)) + '</td></tr>');
      }
      $('#b-table tbody').innerHTML = rows.join('');
    }
    $('#b-csv').addEventListener('click', function () {
      var price = num('b-price'), target = num('b-target') / 100, boost = num('b-boost') / 100;
      var out = [['Keyword conversion rate', 'Max bid (ceiling)', 'Base bid with ' + pct(boost, 0) + ' top-of-search boost']];
      for (var p = 1; p <= 25; p++) { var m = price * target * p / 100; out.push([p + '%', m.toFixed(2), (m / (1 + boost)).toFixed(2)]); }
      csv('dr-ppc-3-number-bid.csv', out);
    });
    live(['b-price', 'b-cogs', 'b-freight', 'b-ref', 'b-fba', 'b-storage', 'b-returns', 'b-target', 'b-cr', 'b-boost'], draw);
  }

  /* ---------- 3. The Bid Stack ---------- */
  // Max click cost = base × (1 + placement boost) × dynamic. Up and down: up to +100% top of search, +50% elsewhere.
  function stack() {
    if (!$('#t-stack')) return;
    var places = [['tos', 'Top of search'], ['ros', 'Rest of search'], ['pp', 'Product pages']];
    function draw() {
      var base = num('s-base'), dyn = $('#s-dyn').value;
      $('#s-table tbody').innerHTML = places.map(function (pl) {
        var b = num('s-' + pl[0]) / 100, up = dyn === 'updown' ? (pl[0] === 'tos' ? 2 : 1.5) : 1;
        var cost = base * (1 + b) * up;
        return '<tr><td>' + pl[1] + '</td><td class="n">' + money(base) + '</td><td class="n">× ' + (1 + b).toFixed(2) + '</td><td class="n">× ' + up.toFixed(1) + '</td><td class="n"><b>' + money(cost) + '</b></td></tr>';
      }).join('');
      $('#s-note').textContent = dyn === 'updown'
        ? 'Up and down: Amazon may raise the bid up to 100% on top of search and up to 50% on other placements, after your placement boost.'
        : dyn === 'down' ? 'Down only: Amazon can lower the bid when a sale looks less likely, never raise it. The figures above are the most a click can cost.'
        : 'Fixed bids: Amazon uses exactly this bid. You pay the auction price, which is at or under it.';
      // Part 2: price each placement off its own conversion rate.
      var cr = places.map(function (pl) { return num('s-cr-' + pl[0]); });
      var floor = Math.min.apply(null, cr.filter(function (x) { return x > 0; }));
      $('#s-price tbody').innerHTML = places.map(function (pl, i) {
        if (!(cr[i] > 0) || !isFinite(floor)) return '<tr><td>' + pl[1] + '</td><td class="n">–</td><td class="n">–</td><td class="n">–</td></tr>';
        var ratio = cr[i] / floor, boost = Math.min((ratio - 1) * 100, 900);
        return '<tr' + (cr[i] === floor ? '' : ' class="hit"') + '><td>' + pl[1] + '</td><td class="n">' + cr[i] + '%</td><td class="n">' + ratio.toFixed(2) + '×</td><td class="n"><b>+' + Math.round(boost) + '%</b>' + (ratio - 1 > 9 ? ' (900% cap)' : '') + '</td></tr>';
      }).join('');
    }
    live(['s-base', 's-tos', 's-ros', 's-pp', 's-dyn', 's-cr-tos', 's-cr-ros', 's-cr-pp'], draw);
  }

  /* ---------- 4. The Two-Lane Sorter ---------- */
  function lanes() {
    if (!$('#t-lanes')) return;
    var data = null;
    var norm = function (s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ''); };
    var n = function (v) { var x = parseFloat(String(v == null ? '' : v).replace(/[$,%\s£€]/g, '')); return isFinite(x) ? x : 0; };
    function pick(head, tests) {
      for (var t = 0; t < tests.length; t++) for (var i = 0; i < head.length; i++) if (tests[t].test(String(head[i]).trim())) return i;
      return -1;
    }
    function load(rows) {
      var h = -1;
      for (var r = 0; r < Math.min(rows.length, 15); r++) if (rows[r].some(function (c) { return /search term|customer search/i.test(String(c)); })) { h = r; break; }
      if (h < 0) { $('#l-status').textContent = 'Couldn’t find a “Customer Search Term” column. Use the Sponsored Products search term report.'; return; }
      var head = rows[h];
      var col = {
        term: pick(head, [/^customer search term$/i, /search term/i]),
        spend: pick(head, [/^spend$/i, /^cost$/i, /^total cost$/i, /spend/i]),
        sales: pick(head, [/total sales/i, /^sales$/i, /^\d+ day sales/i, /sales/i]),
        orders: pick(head, [/total orders/i, /^orders$/i, /^purchases$/i, /orders/i]),
        clicks: pick(head, [/^clicks$/i])
      };
      if (col.spend < 0 || col.sales < 0) { $('#l-status').textContent = 'Found the search terms, but not the Spend and Sales columns.'; return; }
      var terms = {};
      rows.slice(h + 1).forEach(function (r) {
        var t = String(r[col.term] || '').trim(); if (!t) return;
        var k = t.toLowerCase(), e = terms[k] || (terms[k] = { term: t, spend: 0, sales: 0, orders: 0, clicks: 0 });
        e.spend += n(r[col.spend]); e.sales += n(r[col.sales]);
        if (col.orders >= 0) e.orders += n(r[col.orders]);
        if (col.clicks >= 0) e.clicks += n(r[col.clicks]);
      });
      data = Object.keys(terms).map(function (k) { return terms[k]; });
      $('#l-status').textContent = 'Loaded ' + data.length.toLocaleString() + ' search terms. Your file stayed in your browser.';
      draw();
    }
    function readFile(f) {
      if (!f) return;
      $('#l-status').textContent = 'Reading ' + f.name + '…';
      if (/\.xlsx?$/i.test(f.name)) {
        if (!window.XLSX) { $('#l-status').textContent = 'The Excel reader didn’t load. Save the report as CSV and try again.'; return; }
        f.arrayBuffer().then(function (buf) {
          var wb = XLSX.read(buf, { type: 'array' });
          load(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' }));
        });
      } else {
        f.text().then(function (t) { load(parseCSV(t)); });
      }
    }
    function parseCSV(t) {
      var rows = [], row = [], cell = '', q = false, sep = (t.split('\n')[0].match(/\t/g) || []).length > (t.split('\n')[0].match(/,/g) || []).length ? '\t' : ',';
      for (var i = 0; i < t.length; i++) {
        var ch = t[i];
        if (q) { if (ch === '"') { if (t[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
        else if (ch === '"') q = true;
        else if (ch === sep) { row.push(cell); cell = ''; }
        else if (ch === '\n' || ch === '\r') { if (ch === '\r' && t[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
        else cell += ch;
      }
      if (cell || row.length) { row.push(cell); rows.push(row); }
      return rows;
    }
    function draw() {
      var brands = $('#l-brands').value.split(/[,\n]/).map(norm).filter(function (b) { return b.length >= 3; });
      if (!data) return;
      if (!brands.length) { $('#l-out').innerHTML = '<p class="note">Add your brand name above to split the lanes.</p>'; return; }
      var L = { b: { spend: 0, sales: 0, orders: 0, n: 0, list: [] }, nb: { spend: 0, sales: 0, orders: 0, n: 0, list: [] } };
      data.forEach(function (e) {
        var t = norm(e.term), isB = brands.some(function (b) { return t.indexOf(b) >= 0; });
        e.lane = isB ? 'Branded' : 'Non-branded';
        var l = L[isB ? 'b' : 'nb']; l.spend += e.spend; l.sales += e.sales; l.orders += e.orders; l.n++; l.list.push(e);
      });
      var acos = function (l) { return l.sales > 0 ? l.spend / l.sales : null; };
      var all = { spend: L.b.spend + L.nb.spend, sales: L.b.sales + L.nb.sales };
      var card = function (name, l, cls) {
        var a = acos(l);
        return '<div><h3>' + name + '</h3><b class="' + cls + '">' + (a == null ? '–' : pct(a, 1)) + '</b><p>ACoS · ' + money0(l.spend) + ' spend → ' + money0(l.sales) + ' sales' + (l.n != null ? ' · ' + l.n.toLocaleString() + ' terms' : '') + '</p></div>';
      };
      var aNB = acos(L.nb), aAll = all.sales > 0 ? all.spend / all.sales : null;
      var top = function (l) { return l.list.sort(function (x, y) { return y.spend - x.spend; }).slice(0, 8).map(function (e) { return '<tr><td>' + esc(e.term) + '</td><td class="n">' + money(e.spend) + '</td><td class="n">' + money(e.sales) + '</td><td class="n">' + (e.sales > 0 ? pct(e.spend / e.sales, 0) : 'no sales') + '</td></tr>'; }).join(''); };
      $('#l-out').innerHTML = '<div class="lane">' + card('Branded', L.b, 'acos') + card('Non-branded', L.nb, aNB != null && aAll != null && aNB > aAll ? 'bad' : 'acos') + card('Blended (one number)', { spend: all.spend, sales: all.sales, n: null }, '') + '</div>' +
        (aNB != null && aAll != null && aNB > aAll * 1.2 ? '<p class="note"><b style="color:var(--amber)">Branded is covering for it.</b> The blended ' + pct(aAll, 1) + ' looks fine, but your growth lane runs at ' + pct(aNB, 1) + '.</p>' : '') +
        '<div class="row" style="align-items:flex-start;gap:16px"><div style="flex:1;min-width:260px"><h3 style="font-size:14px;margin:16px 0 6px">Top branded terms by spend</h3><div class="tbl"><table><thead><tr><th>Term</th><th class="n">Spend</th><th class="n">Sales</th><th class="n">ACoS</th></tr></thead><tbody>' + top(L.b) + '</tbody></table></div></div>' +
        '<div style="flex:1;min-width:260px"><h3 style="font-size:14px;margin:16px 0 6px">Top non-branded terms by spend</h3><div class="tbl"><table><thead><tr><th>Term</th><th class="n">Spend</th><th class="n">Sales</th><th class="n">ACoS</th></tr></thead><tbody>' + top(L.nb) + '</tbody></table></div></div></div>' +
        '<div class="row"><button class="btn2" type="button" id="l-csv">Download every term with its lane (CSV)</button></div>';
      $('#l-csv').addEventListener('click', function () {
        csv('dr-ppc-two-lane-sort.csv', [['Search term', 'Lane', 'Spend', 'Sales', 'Orders', 'Clicks', 'ACoS']].concat(data.map(function (e) { return [e.term, e.lane, e.spend.toFixed(2), e.sales.toFixed(2), e.orders, e.clicks, e.sales > 0 ? (e.spend / e.sales * 100).toFixed(1) + '%' : '']; })));
      });
    }
    var drop = $('#l-drop');
    ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('on'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('on'); }); });
    drop.addEventListener('drop', function (e) { readFile(e.dataTransfer.files[0]); });
    $('#l-file').addEventListener('change', function (e) { readFile(e.target.files[0]); });
    $('#l-brands').addEventListener('input', draw);
    // Quick math, no file needed.
    live(['q-bs', 'q-bv', 'q-ns', 'q-nv'], function () {
      var b = { spend: num('q-bs'), sales: num('q-bv') }, nb = { spend: num('q-ns'), sales: num('q-nv') }, all = { spend: b.spend + nb.spend, sales: b.sales + nb.sales };
      var a = function (l) { return l.sales > 0 ? pct(l.spend / l.sales, 1) : '–'; };
      $('#q-out').innerHTML = '<div><b>' + a(b) + '</b><span>branded ACoS</span></div><div class="' + (nb.sales && all.sales && nb.spend / nb.sales > all.spend / all.sales ? 'warn' : '') + '"><b>' + a(nb) + '</b><span>non-branded ACoS: your growth lane</span></div><div><b>' + a(all) + '</b><span>blended: ' + money0(all.spend) + ' on ' + money0(all.sales) + '</span></div>';
    });
  }

  /* ---------- 5. The Date Windows ---------- */
  function lag() {
    if (!$('#t-lag')) return;
    var DAY = 864e5;
    var iso = function (d) { return d.toISOString().slice(0, 10); };
    var nice = function (d) { return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' }); };
    var add = function (d, n) { return new Date(d.getTime() + n * DAY); };
    var parse = function (v) { var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || ''); return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null; };
    var t = new Date(); $('#g-today').value = iso(new Date(Date.UTC(t.getFullYear(), t.getMonth(), t.getDate())));
    function draw() {
      var today = parse($('#g-today').value); if (!today) return;
      var type = $('#g-type').value, credit = type === 'sp' ? 7 : 14, settle = type === 'sp' ? 3 : 7;
      var end = add(today, -settle), start = add(end, -29);
      var out = '<div><b>' + nice(start) + ' – ' + nice(end) + '</b><span>judge bids, keywords and search terms on these 30 days (they end ' + settle + ' days ago)</span></div>' +
        '<div class="warn"><b>' + nice(add(today, -settle + 1)) + ' – ' + nice(add(today, -1)) + '</b><span>still filling in: late orders are still landing on these days. Don’t cut on them.</span></div>' +
        '<div><b>' + nice(add(today, -1)) + '</b><span>yesterday is fine for spend and budget checks. Just not for sales or ACoS.</span></div>';
      var ch = parse($('#g-change').value);
      if (ch) {
        var before0 = add(ch, -credit), before1 = add(ch, -1), after1 = add(ch, credit - 1), judge = add(after1, settle);
        out += '<div><b>' + nice(judge) + '</b><span>judge your change on this day: compare ' + nice(before0) + '–' + nice(before1) + ' (before) with ' + nice(ch) + '–' + nice(after1) + ' (after)</span></div>';
      }
      $('#g-out').innerHTML = out;
    }
    live(['g-today', 'g-type', 'g-change'], draw);
    $('#g-csv').addEventListener('click', function () {
      csv('dr-ppc-change-log.csv', [['Date', 'Campaign', 'What changed', 'From', 'To', 'Why', 'Judge on'], ['', '', 'e.g. exact bid on "insulated water bottle"', '$1.20', '$1.44', 'Past target for 30 days ending 3 days ago', '']]);
    });
  }

  document.addEventListener('DOMContentLoaded', function () { rule(); bid(); stack(); lanes(); lag(); });
})();
