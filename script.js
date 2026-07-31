/* =============================================================
   Outward/Dispatch — script.js
   - In-browser Excel/CSV ingestion via SheetJS (no uploads)
   - Live Google Sheet sync (public CSV export + gviz fallback)
   - Data-cleaning layer for known data-entry typos
   - IndexedDB persistence (reload-safe, zero backend)
   - KPI grid, Chart.js visuals, interactive sortable matrix
   ============================================================= */
'use strict';

(function () {
  /* -------------------------------------------------------------
     Config
     ------------------------------------------------------------- */
  var SHEET_ID = '1C5JVg39Ad7gcdAJCb6qUO6qBadhm-umqRwD8CFt5oc0';
  var SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?format=csv';
  var SHEET_GVIZ_URL = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/gviz/tq?tqx=out:json&headers=1';

  var DB_NAME = 'outward-dashboard';
  var DB_VERSION = 1;
  var STORE = 'records';
  var META_STORE = 'meta';

  var REQUIRED_COLUMNS = ['Order Date', 'Customer Name', 'Branch', 'Location', 'Priority', 'Dispatch date', 'Ack'];

  /* Header aliases → canonical columns. Longest/most specific keys first
     so prefix-matching in the gviz fallback resolves correctly. */
  var HEADER_ALIASES = {
    'order date': 'Order Date',
    'orderdate': 'Order Date',
    'order no': 'Order Date',
    'order': 'Order Date',
    'od': 'Order Date',
    'date': 'Order Date',
    'customer name': 'Customer Name',
    'customername': 'Customer Name',
    'customer': 'Customer Name',
    'client name': 'Customer Name',
    'client': 'Customer Name',
    'branch': 'Branch',
    'branch name': 'Branch',
    'location': 'Location',
    'priority': 'Priority',
    'priority level': 'Priority',
    'dispatch date': 'Dispatch date',
    'dispatchdate': 'Dispatch date',
    'dispatch': 'Dispatch date',
    'dispatched date': 'Dispatch date',
    'despatch date': 'Dispatch date',
    'ship date': 'Dispatch date',
    'acknowledgement status': 'Ack',
    'acknowledgment status': 'Ack',
    'acknowledgement': 'Ack',
    'acknowledgment': 'Ack',
    'ack status': 'Ack',
    'ack': 'Ack'
  };

  /* Known data-entry typos → canonical place names (Hyderabad area) */
  var PLACE_TYPO_MAP = {
    'gcachibowli': 'Gachibowli',
    'gachibowli': 'Gachibowli',
    'kkatpally': 'Kukatpally',
    'kukatpally': 'Kukatpally',
    'malkpet': 'Malakpet',
    'malakpet': 'Malakpet',
    'myapur': 'Miyapur',
    'miyapur': 'Miyapur',
    'mehndipatnam': 'Mehdipatnam',
    'mehdipatnam': 'Mehdipatnam',
    'chanda nagar': 'Chandanagar',
    'chandanagar': 'Chandanagar',
    'chandanagr': 'Chandanagar',
    'hitech-city': 'Hitech City',
    'hitech city': 'Hitech City',
    'hitechcity': 'Hitech City',
    'jubileehills': 'Jubilee Hills',
    'jubilee hills': 'Jubilee Hills',
    'jublieehills': 'Jubilee Hills',
    'banjarahills': 'Banjara Hills',
    'banjara hills': 'Banjara Hills',
    'banjarahils': 'Banjara Hills',
    'panjagutta': 'Punjagutta',
    'punjagutta': 'Punjagutta',
    'lb-nagar': 'LB Nagar',
    'lb nagar': 'LB Nagar',
    'lbnagar': 'LB Nagar',
    'sr-nagar': 'SR Nagar',
    'sr nagar': 'SR Nagar',
    'srnagar': 'SR Nagar',
    'dilsukhnagar': 'Dilsukhnagar',
    'dilsuknagar': 'Dilsukhnagar',
    'himayatnagar': 'Himayatnagar',
    'himayathnagar': 'Himayatnagar',
    'kompally': 'Kompally',
    'kompalli': 'Kompally',
    'tarnaka': 'Tarnaka',
    'nizampet': 'Nizampet',
    'secunderabad': 'Secunderabad',
    'secendrabad': 'Secunderabad',
    'ameerpet': 'Ameerpet',
    'ameerpeth': 'Ameerpet',
    'begumpet': 'Begumpet',
    'kondapur': 'Kondapur',
    'madhapur': 'Madhapur',
    'madapur': 'Madhapur',
    'uppal': 'Uppal',
    'abids': 'Abids'
  };

  var PRIORITY_MAP = {
    'p1': 'P1', 'priority1': 'P1', 'high': 'P1', 'urgent': 'P1', 'top': 'P1', '1': 'P1',
    'p2': 'P2', 'priority2': 'P2', 'medium': 'P2', 'normal': 'P2', '2': 'P2',
    'p3': 'P3', 'priority3': 'P3', 'low': 'P3', '3': 'P3',
    'p4': 'P4', 'priority4': 'P4', 'none': 'P4', '4': 'P4'
  };

  var ACK_DONE = ['done', 'completed', 'complete', 'delivered', 'acked', 'acknowledged', 'ok', 'yes'];
  var ACK_TRANSIT = ['in transit', 'intransit', 'in-transit', 'transit', 'shipped', 'dispatched', 'on the way', 'inroute', 'in route', 'onroute', 'on route'];

  var PRIORITY_RANK = { 'P1': 0, 'P2': 1, 'P3': 2, 'P4': 3, '—': 4 };
  var ACK_RANK = { 'Done': 0, 'In Transit': 1, 'Pending': 2 };

  /* -------------------------------------------------------------
     State
     ------------------------------------------------------------- */
  var state = {
    records: [],
    source: null,
    priority: 'ALL',
    ack: 'ALL',
    query: '',
    sortKey: 'orderDate',
    sortDir: 'desc'
  };

  var db = null;
  var barChart = null;
  var doughnutChart = null;
  var syncing = false;

  /* -------------------------------------------------------------
     Data cleaning layer (pure — unit-testable)
     ------------------------------------------------------------- */
  function cleanPlace(v) {
    var s = String(v == null ? '' : v).trim().replace(/\s+/g, ' ');
    if (!s) return '';
    var key = s.toLowerCase();
    if (PLACE_TYPO_MAP[key]) return PLACE_TYPO_MAP[key];
    return s.split(/[\s-]+/).filter(Boolean).map(function (w) {
      if (w === w.toUpperCase()) return w;
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    }).join(' ');
  }

  function cleanCustomer(v) {
    return String(v == null ? '' : v).trim().replace(/\s+/g, ' ').replace(/\r?\n/g, ' ');
  }

  function normalizePriority(v) {
    var key = String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, '');
    return PRIORITY_MAP[key] || '—';
  }

  function normalizeAck(v) {
    var s = String(v == null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' ');
    if (!s) return 'Pending';
    if (ACK_DONE.indexOf(s) !== -1) return 'Done';
    if (ACK_TRANSIT.indexOf(s) !== -1) return 'In Transit';
    return 'Pending';
  }

  function parseDate(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'number' && isFinite(v)) {
      if (typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
        var d = XLSX.SSF.parse_date_code(v);
        if (d) return new Date(d.y, d.m - 1, d.d);
      }
      return null;
    }
    var s = String(v).trim();
    if (!s) return null;

    var gviz = s.match(/^Date\((\d{4}),(\d{1,2}),(\d{1,2})\)$/);
    if (gviz) return new Date(+gviz[1], +gviz[2], +gviz[3]);

    /* dd/mm/yyyy (dataset format) — defensively swap if month>12 */
    var sl = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$/);
    if (sl) {
      var day = +sl[1], mon = +sl[2], yr = +sl[3];
      if (yr < 100) yr += 2000;
      if (mon > 12 && day <= 12) {
        var t = day; day = mon; mon = t;
      }
      if (mon >= 1 && mon <= 12 && day >= 1 && day <= 31) {
        var dt = new Date(yr, mon - 1, day);
        return isNaN(dt.getTime()) ? null : dt;
      }
      return null;
    }

    var iso = new Date(s);
    return isNaN(iso.getTime()) ? null : iso;
  }

  function normalizeRecord(raw) {
    return {
      id: makeId(),
      orderDate: parseDate(raw['Order Date']),
      customer: cleanCustomer(raw['Customer Name']),
      branch: cleanPlace(raw['Branch']),
      location: cleanPlace(raw['Location']),
      priority: normalizePriority(raw['Priority']),
      dispatchDate: parseDate(raw['Dispatch date']),
      ack: normalizeAck(raw['Ack'])
    };
  }

  function makeId() {
    if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
    return 'r-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 10);
  }

  /* -------------------------------------------------------------
     Workbook → records (header-tolerant, error-bounded)
     ------------------------------------------------------------- */
  function normHeaderKey(v) {
    return String(v == null ? '' : v)
      .toLowerCase()
      .replace(/[-_.]/g, ' ')
      .replace(/[()[\]]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function detectHeaders(rows) {
    var headerMap = {};
    var startIndex = 0;
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row || !row.some(function (c) { return String(c == null ? '' : c).trim() !== ''; })) {
        startIndex = i + 1;
        continue;
      }
      row.forEach(function (cell, colIdx) {
        var key = normHeaderKey(cell);
        if (HEADER_ALIASES[key] && !(HEADER_ALIASES[key] in headerMap)) {
          headerMap[HEADER_ALIASES[key]] = colIdx;
        }
      });
      startIndex = i + 1;
      break;
    }
    var missing = REQUIRED_COLUMNS.filter(function (c) { return !(c in headerMap); });
    return { headerMap: headerMap, missing: missing, startIndex: startIndex };
  }

  function buildRecords(rows, headerMap, startIndex) {
    var records = [];
    for (var i = startIndex; i < rows.length; i++) {
      var row = rows[i];
      if (!row || !row.some(function (c) { return String(c == null ? '' : c).trim() !== ''; })) continue;
      var raw = {};
      for (var canon in headerMap) raw[canon] = row[headerMap[canon]] == null ? '' : row[headerMap[canon]];
      var rec = normalizeRecord(raw);
      if (!rec.customer && !rec.branch && !rec.location) continue;
      records.push(rec);
    }
    return records;
  }

  function ingestWorkbook(wb) {
    var allRows = [];
    for (var i = 0; i < wb.SheetNames.length; i++) {
      var sheet = wb.Sheets[wb.SheetNames[i]];
      if (!sheet || !sheet['!ref']) continue;
      allRows = allRows.concat(XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true }));
    }
    var detection = detectHeaders(allRows);
    if (detection.missing.length) return { records: [], missing: detection.missing };
    return { records: buildRecords(allRows, detection.headerMap, detection.startIndex), missing: [] };
  }

  /* gviz JSON fallback for the live sheet (best effort) */
  function parseGvizRows(text) {
    var json;
    try {
      var body = String(text).trim();
      body = body.replace(/^\/\*O_o\*\//, '');
      body = body.slice(body.indexOf('(') + 1);
      body = body.replace(/\);?\s*$/, '');
      json = JSON.parse(body);
    } catch (e) {
      throw new Error('Could not parse the Google Sheet response.');
    }
    if (!json || !json.table) throw new Error('Google Sheet returned no table.');
    var cols = json.table.cols || [];
    var rows = json.table.rows || [];
    var headerIdx = {};
    var aliasKeys = Object.keys(HEADER_ALIASES);
    cols.forEach(function (col, i) {
      var label = col && col.label ? String(col.label) : '';
      var norm = label.toLowerCase();
      var canonical = HEADER_ALIASES[norm];
      if (!canonical) {
        for (var a = 0; a < aliasKeys.length; a++) {
          if (norm.indexOf(aliasKeys[a]) === 0) {
            canonical = HEADER_ALIASES[aliasKeys[a]];
            break;
          }
        }
      }
      if (canonical && !(canonical in headerIdx)) headerIdx[canonical] = i;
    });
    var missing = REQUIRED_COLUMNS.filter(function (c) { return !(c in headerIdx); });
    if (missing.length) throw new Error('Required columns missing: ' + missing.join(', '));
    return rows.filter(function (r) { return r && r.c; }).map(function (r) {
      var raw = {};
      for (var canon in headerIdx) {
        var cell = r.c[headerIdx[canon]];
        var val = '';
        if (cell) {
          if (cell.f != null) val = cell.f;
          else if (Array.isArray(cell.v)) val = cell.v.join(', ');
          else val = cell.v;
        }
        raw[canon] = val == null ? '' : val;
      }
      return raw;
    });
  }

  /* -------------------------------------------------------------
     Derived computations
     ------------------------------------------------------------- */
  function computeKpis(rows) {
    var total = rows.length;
    var bottlenecks = 0, done = 0, transit = 0, pending = 0, days = 0, velocityCount = 0;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r.priority === 'P1' && r.ack !== 'Done') bottlenecks++;
      if (r.ack === 'Done') done++;
      else if (r.ack === 'In Transit') transit++;
      else pending++;
      if (r.orderDate && r.dispatchDate) {
        var diff = (r.dispatchDate.getTime() - r.orderDate.getTime()) / 86400000;
        if (diff >= 0) { days += diff; velocityCount++; }
      }
    }
    return {
      total: total,
      bottlenecks: bottlenecks,
      done: done,
      transit: transit,
      pending: pending,
      velocity: velocityCount ? days / velocityCount : 0,
      velocityCount: velocityCount,
      fulfillment: total ? (done / total) * 100 : 0
    };
  }

  function buildBarData(rows) {
    var counts = {};
    rows.forEach(function (r) {
      var c = r.customer || 'Unknown';
      counts[c] = (counts[c] || 0) + 1;
    });
    var entries = Object.keys(counts).map(function (k) { return [k, counts[k]]; })
      .sort(function (a, b) { return b[1] - a[1]; });
    var top = entries.slice(0, 12);
    var labels = top.map(function (e) { return e[0]; });
    var data = top.map(function (e) { return e[1]; });
    var rest = entries.slice(12).reduce(function (sum, e) { return sum + e[1]; }, 0);
    if (rest > 0) { labels.push('Others'); data.push(rest); }
    return { labels: labels, data: data };
  }

  function buildDoughnutData(rows) {
    var k = computeKpis(rows);
    return { labels: ['Done', 'In Transit', 'Pending'], data: [k.done, k.transit, k.pending], total: k.total };
  }

  function filteredRows() {
    var rows = state.records;
    if (state.priority !== 'ALL') rows = rows.filter(function (r) { return r.priority === state.priority; });
    if (state.ack !== 'ALL') rows = rows.filter(function (r) { return r.ack === state.ack; });
    if (state.query) {
      var q = state.query.toLowerCase();
      rows = rows.filter(function (r) {
        return r.customer.toLowerCase().indexOf(q) !== -1 ||
          r.branch.toLowerCase().indexOf(q) !== -1 ||
          r.location.toLowerCase().indexOf(q) !== -1;
      });
    }
    return rows;
  }

  function sortRows(rows, key, dir) {
    var mul = dir === 'asc' ? 1 : -1;
    return rows.slice().sort(function (a, b) {
      var cmp = 0;
      if (key === 'orderDate' || key === 'dispatchDate') {
        var av = a[key] ? a[key].getTime() : Infinity;
        var bv = b[key] ? b[key].getTime() : Infinity;
        cmp = av === bv ? 0 : (av < bv ? -1 : 1);
      } else if (key === 'priority') {
        cmp = (PRIORITY_RANK[a.priority] != null ? PRIORITY_RANK[a.priority] : 4) -
              (PRIORITY_RANK[b.priority] != null ? PRIORITY_RANK[b.priority] : 4);
      } else if (key === 'ack') {
        cmp = (ACK_RANK[a.ack] != null ? ACK_RANK[a.ack] : 2) -
              (ACK_RANK[b.ack] != null ? ACK_RANK[b.ack] : 2);
      } else {
        cmp = String(a[key] == null ? '' : a[key]).localeCompare(
          String(b[key] == null ? '' : b[key]), undefined, { numeric: true, sensitivity: 'base' });
      }
      return cmp * mul;
    });
  }

  /* -------------------------------------------------------------
     IndexedDB persistence
     ------------------------------------------------------------- */
  function openDB() {
    return new Promise(function (resolve) {
      if (typeof indexedDB === 'undefined') { resolve(null); return; }
      try {
        var req = indexedDB.open(DB_NAME, DB_VERSION);
        req.onupgradeneeded = function (e) {
          var d = e.target.result;
          if (!d.objectStoreNames.contains(STORE)) d.createObjectStore(STORE, { keyPath: 'id' });
          if (!d.objectStoreNames.contains(META_STORE)) d.createObjectStore(META_STORE, { keyPath: 'key' });
        };
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { console.warn('IndexedDB open failed:', req.error); resolve(null); };
        req.onblocked = function () { resolve(null); };
      } catch (e) {
        console.warn('IndexedDB unavailable:', e);
        resolve(null);
      }
    });
  }

  function txDone(tx) {
    return new Promise(function (resolve, reject) {
      tx.oncomplete = resolve;
      tx.onerror = function () { reject(tx.error); };
      tx.onabort = function () { reject(tx.error); };
    });
  }

  function persistRecords(records, source) {
    if (!db) return Promise.resolve();
    return new Promise(function (resolve, reject) {
      try {
        var tx = db.transaction([STORE, META_STORE], 'readwrite');
        tx.objectStore(STORE).clear();
        for (var i = 0; i < records.length; i++) tx.objectStore(STORE).put(records[i]);
        tx.objectStore(META_STORE).put({
          key: 'source',
          name: source ? source.name : null,
          syncedAt: source && source.syncedAt ? source.syncedAt : Date.now()
        });
        txDone(tx).then(resolve, reject);
      } catch (e) {
        console.warn('IndexedDB write failed:', e);
        reject(e);
      }
    });
  }

  function loadPersisted() {
    if (!db) return Promise.resolve(null);
    return new Promise(function (resolve) {
      try {
        var tx = db.transaction([STORE, META_STORE], 'readonly');
        var recsReq = tx.objectStore(STORE).getAll();
        var metaReq = tx.objectStore(META_STORE).get('source');
        txDone(tx).then(function () {
          resolve({ records: recsReq.result || [], meta: metaReq.result || null });
        }, function (e) {
          console.warn('IndexedDB read failed:', e);
          resolve(null);
        });
      } catch (e) {
        console.warn('IndexedDB read failed:', e);
        resolve(null);
      }
    });
  }

  /* -------------------------------------------------------------
     Rendering helpers
     ------------------------------------------------------------- */
  function $(id) { return document.getElementById(id); }

  function setText(id, value) {
    var el = $(id);
    if (el) el.textContent = value;
  }

  function formatNum(n) { return Number(n || 0).toLocaleString('en-IN'); }

  function formatDate(d) {
    if (!d) return '<span class="pending-date">—</span>';
    var dd = String(d.getDate()).padStart(2, '0');
    var mm = String(d.getMonth() + 1).padStart(2, '0');
    return dd + '/' + mm + '/' + d.getFullYear();
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function isFileProtocol() {
    return typeof location !== 'undefined' && location.protocol === 'file:';
  }

  function refreshIcons() {
    if (typeof lucide !== 'undefined' && lucide.createIcons) {
      try { lucide.createIcons(); } catch (e) { /* no-op */ }
    }
  }

  function setChartEmpty(id, show) {
    var el = $(id);
    if (el) el.hidden = !show;
    refreshIcons();
  }

  function toast(message, type, timeout) {
    type = type || 'info';
    timeout = timeout || 4200;
    var stack = $('toastStack');
    if (!stack) return;
    var el = document.createElement('div');
    el.className = 'toast toast-' + type;
    var icon = type === 'success' ? 'check-circle-2' : (type === 'error' ? 'alert-triangle' : 'info');
    el.innerHTML = '<i data-lucide="' + icon + '"></i><span></span>';
    el.querySelector('span').textContent = message;
    stack.appendChild(el);
    refreshIcons();
    setTimeout(function () {
      el.classList.add('toast-out');
      setTimeout(function () { if (el.parentNode) el.parentNode.removeChild(el); }, 220);
    }, timeout);
  }

  function showError(message) {
    var banner = $('errorBanner');
    if (banner) {
      banner.innerHTML =
        '<i data-lucide="alert-triangle"></i>' +
        '<span></span>' +
        '<button type="button" class="error-close" aria-label="Dismiss"><i data-lucide="x"></i></button>';
      banner.querySelector('span').textContent = message;
      banner.hidden = false;
      banner.querySelector('.error-close').addEventListener('click', function () { banner.hidden = true; });
      refreshIcons();
      if (banner.scrollIntoView) banner.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    toast(message, 'error');
  }

  var LOCAL_NOTICE_MSG =
    'Google Sheets sync is blocked when this page is opened directly from the file system. ' +
    'Run "npx serve" (or "python -m http.server") in this folder and open it via http:// — ' +
    'or simply drag & drop the Excel/CSV export here.';

  function showLocalNotice() {
    var el = $('localNotice');
    if (!el || el.hidden === false) return;
    var span = el.querySelector('span');
    if (span) span.textContent = LOCAL_NOTICE_MSG;
    el.hidden = false;
    refreshIcons();
    var close = el.querySelector('.error-close');
    if (close) close.addEventListener('click', function () { el.hidden = true; });
  }

  function describeSyncError(err) {
    if (isFileProtocol()) {
      return 'Google Sheets blocks data access for pages opened from the file system. ' +
        'Serve this folder over HTTP ("npx serve" or "python -m http.server"), then reopen — ' +
        'or drag & drop the Excel/CSV export directly into the drop zone.';
    }
    if (err && err.message === 'Failed to fetch') {
      return 'Could not reach Google Sheets — usually a network issue or an ad/tracker blocker. ' +
        'Check your connection, disable browser extensions for this page, and try again.';
    }
    return err && err.message ? err.message : 'Unknown error.';
  }

  function priorityBadge(p) {
    var clsMap = { P1: 'badge-p1', P2: 'badge-p2', P3: 'badge-p3', P4: 'badge-p4' };
    var dotMap = { P1: 'dot-p1', P2: 'dot-p2', P3: 'dot-p3', P4: 'dot-p4' };
    var cls = clsMap[p] || 'badge-neutral';
    var dot = dotMap[p] || 'dot-p4';
    return '<span class="badge ' + cls + '"><span class="dot ' + dot + '"></span>' + escapeHtml(p) + '</span>';
  }

  function ackBadge(a) {
    var clsMap = { 'Done': 'badge-done', 'In Transit': 'badge-transit', 'Pending': 'badge-pending' };
    var dotMap = { 'Done': 'dot-done', 'In Transit': 'dot-transit', 'Pending': 'dot-pending' };
    var cls = clsMap[a] || 'badge-neutral';
    var dot = dotMap[a] || 'dot-pending';
    return '<span class="badge ' + cls + '"><span class="dot ' + dot + '"></span>' + escapeHtml(a) + '</span>';
  }

  /* -------------------------------------------------------------
     KPI grid
     ------------------------------------------------------------- */
  function renderKpis(rows) {
    var k = computeKpis(rows);
    var scope = rows.length === state.records.length;
    setText('kpiVolume', formatNum(k.total));
    setText('kpiVolumeNote', scope ? 'records in scope' : formatNum(k.total) + ' of ' + formatNum(state.records.length) + ' in scope');
    setText('kpiBottleneck', formatNum(k.bottlenecks));
    setText('kpiVelocity', k.velocityCount ? k.velocity.toFixed(1) + 'd' : '—');
    setText('kpiFulfillment', k.total ? k.fulfillment.toFixed(1) + '%' : '—');
  }

  /* -------------------------------------------------------------
     Charts
     ------------------------------------------------------------- */
  var centerTextPlugin = {
    id: 'centerText',
    afterDraw: function (chart, args, opts) {
      if (!opts || !opts.total || !chart.data || !chart.data.datasets.length) return;
      var dataset = chart.data.datasets[0];
      if (!dataset || !dataset.data || !dataset.data.length) return;
      var meta = chart.getDatasetMeta(0);
      if (!meta.data || !meta.data.length) return;
      var point = meta.data[0];
      var x = point ? point.x : chart.width / 2;
      var y = point ? point.y : chart.height / 2;
      var ctx = chart.ctx;
      ctx.save();
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = "600 24px 'Inter', sans-serif";
      ctx.fillStyle = '#f7f8f8';
      ctx.fillText(Number(opts.total).toLocaleString('en-IN'), x, y - 7);
      ctx.font = "400 11px 'Inter', sans-serif";
      ctx.fillStyle = '#8a8f98';
      ctx.fillText('records', x, y + 12);
      ctx.restore();
    }
  };

  function createCharts() {
    if (typeof Chart === 'undefined') {
      showError('Chart.js failed to load — check your internet connection and reload.');
      return;
    }
    Chart.defaults.font.family = "'Inter', system-ui, sans-serif";
    Chart.defaults.font.size = 12;
    Chart.defaults.animation.duration = 350;

    var barCtx = $('barChart');
    if (barCtx) {
      barChart = new Chart(barCtx, {
        type: 'bar',
        data: {
          labels: [],
          datasets: [{
            data: [],
            backgroundColor: '#5e6ad2',
            hoverBackgroundColor: '#828fff',
            borderRadius: 5,
            borderSkipped: false,
            maxBarThickness: 20
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              backgroundColor: '#18191a',
              borderColor: '#34343a',
              borderWidth: 1,
              titleColor: '#f7f8f8',
              bodyColor: '#d0d6e0',
              padding: 10,
              cornerRadius: 8,
              displayColors: false
            }
          },
          scales: {
            x: {
              beginAtZero: true,
              grid: { color: '#23252a' },
              border: { color: '#23252a' },
              ticks: { color: '#8a8f98', precision: 0, callback: function (v) { return Number(v).toLocaleString('en-IN'); } }
            },
            y: {
              grid: { display: false },
              border: { color: '#23252a' },
              ticks: { color: '#d0d6e0', autoSkip: false, font: { size: 12 } }
            }
          }
        }
      });
    }

    var doughnutCtx = $('doughnutChart');
    if (doughnutCtx) {
      doughnutChart = new Chart(doughnutCtx, {
        type: 'doughnut',
        data: {
          labels: ['Done', 'In Transit', 'Pending'],
          datasets: [{
            data: [],
            backgroundColor: ['#27a644', '#4aa8ff', '#f59e0b'],
            borderColor: '#0f1011',
            borderWidth: 4,
            hoverOffset: 8
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          cutout: '68%',
          plugins: {
            centerText: { total: 0 },
            legend: {
              position: 'bottom',
              labels: { color: '#d0d6e0', usePointStyle: true, pointStyle: 'circle', boxWidth: 8, padding: 18 }
            },
            tooltip: {
              backgroundColor: '#18191a',
              borderColor: '#34343a',
              borderWidth: 1,
              titleColor: '#f7f8f8',
              bodyColor: '#d0d6e0',
              padding: 10,
              cornerRadius: 8
            }
          }
        },
        plugins: [centerTextPlugin]
      });
    }
  }

  function renderCharts(rows) {
    var barData = buildBarData(rows);
    setText('barMeta', rows.length ? barData.labels.length + ' segments' : '—');
    setChartEmpty('barEmpty', rows.length === 0);
    if (barChart) {
      var isOthers = function (label, idx) { return label === 'Others' && idx === barData.labels.length - 1; };
      var colors = barData.labels.map(function (_, idx) {
        return isOthers(barData.labels[idx], idx) ? 'rgba(94,106,210,0.45)' : '#5e6ad2';
      });
      var hoverColors = barData.labels.map(function (_, idx) {
        return isOthers(barData.labels[idx], idx) ? 'rgba(130,143,255,0.55)' : '#828fff';
      });
      barChart.data.labels = barData.labels;
      barChart.data.datasets[0].data = barData.data;
      barChart.data.datasets[0].backgroundColor = colors;
      barChart.data.datasets[0].hoverBackgroundColor = hoverColors;
      barChart.update();
    }

    var dData = buildDoughnutData(rows);
    setText('doughnutMeta', rows.length ? formatNum(dData.total) + ' records' : '—');
    setChartEmpty('doughnutEmpty', rows.length === 0);
    if (doughnutChart) {
      doughnutChart.data.datasets[0].data = dData.data;
      doughnutChart.options.plugins.centerText.total = dData.total;
      doughnutChart.update();
    }
  }

  /* -------------------------------------------------------------
     Matrix table
     ------------------------------------------------------------- */
  function renderSortIcons() {
    var headers = document.querySelectorAll('th.sortable');
    for (var i = 0; i < headers.length; i++) {
      var th = headers[i];
      var ind = th.querySelector('.sort-ind');
      var active = th.getAttribute('data-key') === state.sortKey;
      th.classList.toggle('sorted', active);
      th.setAttribute('aria-sort', active ? (state.sortDir === 'asc' ? 'ascending' : 'descending') : 'none');
      if (ind) {
        ind.innerHTML = active
          ? '<i data-lucide="' + (state.sortDir === 'asc' ? 'chevron-up' : 'chevron-down') + '"></i>'
          : '<i data-lucide="chevrons-up-down"></i>';
      }
    }
    refreshIcons();
  }

  function renderTable() {
    var tbody = $('tableBody');
    if (!tbody) return;
    var rows = sortRows(filteredRows(), state.sortKey, state.sortDir);
    setText('rowCount', formatNum(rows.length) + ' of ' + formatNum(state.records.length) + ' records');
    if (!rows.length) {
      var msg = state.records.length
        ? 'No records match the current filters'
        : 'No data yet — drag & drop a file above, or sync the Google Sheet';
      tbody.innerHTML = '<tr class="row-empty"><td colspan="7">' + msg + '</td></tr>';
      return;
    }
    var html = '';
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      html +=
        '<tr>' +
        '<td class="td-date">' + formatDate(r.orderDate) + '</td>' +
        '<td class="td-strong">' + escapeHtml(r.customer || '—') + '</td>' +
        '<td>' + escapeHtml(r.branch || '—') + '</td>' +
        '<td>' + escapeHtml(r.location || '—') + '</td>' +
        '<td>' + priorityBadge(r.priority) + '</td>' +
        '<td class="td-date">' + formatDate(r.dispatchDate) + '</td>' +
        '<td>' + ackBadge(r.ack) + '</td>' +
        '</tr>';
    }
    tbody.innerHTML = html;
  }

  /* -------------------------------------------------------------
     Source badge
     ------------------------------------------------------------- */
  function renderSource() {
    var el = $('dataSourceBadge');
    if (!el) return;
    if (!state.records.length) {
      el.textContent = 'No data loaded';
      el.removeAttribute('title');
      return;
    }
    var label = state.source && state.source.name
      ? state.source.name + ' · ' + formatNum(state.records.length) + ' rows'
      : formatNum(state.records.length) + ' rows';
    el.textContent = label;
    if (state.source && state.source.syncedAt) {
      el.title = 'Loaded at ' + new Date(state.source.syncedAt).toLocaleString();
    } else {
      el.removeAttribute('title');
    }
  }

  /* -------------------------------------------------------------
     Filters, search, sort
     ------------------------------------------------------------- */
  function renderAll() {
    var rows = filteredRows();
    renderKpis(rows);
    renderCharts(rows);
    renderTable();
    renderSource();
    refreshIcons();
  }

  function resetFilters(silent) {
    state.priority = 'ALL';
    state.ack = 'ALL';
    state.query = '';
    var search = $('searchInput');
    if (search) search.value = '';
    var priorityPills = document.querySelectorAll('#priorityFilter .pill');
    for (var i = 0; i < priorityPills.length; i++) {
      priorityPills[i].classList.toggle('is-active', priorityPills[i].getAttribute('data-priority') === 'ALL');
    }
    var ackPills = document.querySelectorAll('#ackFilter .pill');
    for (var j = 0; j < ackPills.length; j++) {
      ackPills[j].classList.toggle('is-active', ackPills[j].getAttribute('data-ack') === 'ALL');
    }
    if (!silent) renderAll();
  }

  function setPriorityFilter(p) {
    state.priority = p;
    var pills = document.querySelectorAll('#priorityFilter .pill');
    for (var i = 0; i < pills.length; i++) {
      pills[i].classList.toggle('is-active', pills[i].getAttribute('data-priority') === p);
    }
    renderAll();
  }

  function setAckFilter(a) {
    state.ack = a;
    var pills = document.querySelectorAll('#ackFilter .pill');
    for (var i = 0; i < pills.length; i++) {
      pills[i].classList.toggle('is-active', pills[i].getAttribute('data-ack') === a);
    }
    renderAll();
  }

  /* -------------------------------------------------------------
     Ingestion flows
     ------------------------------------------------------------- */
  function applyRecords(records, source) {
    state.records = records;
    state.source = source || null;
    resetFilters(true);
    renderAll();
    persistRecords(records, source).catch(function (e) {
      console.warn('Could not persist to IndexedDB:', e);
    });
  }

  function handleFile(file) {
    if (!file) return;
    var ext = (file.name.split('.').pop() || '').toLowerCase();
    var ok = ['xlsx', 'xls', 'csv'].indexOf(ext) !== -1 || file.type === 'text/csv';
    if (!ok) {
      showError('"' + file.name + '" is not a supported workbook. Please upload .xlsx, .xls or .csv.');
      return;
    }
    var title = $('dzTitle');
    var originalTitle = title ? title.textContent : '';
    if (title) title.textContent = 'Processing ' + file.name + '…';

    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        if (typeof XLSX === 'undefined') throw new Error('SheetJS failed to load — check your internet connection.');
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        var result = ingestWorkbook(wb);
        if (result.missing.length) {
          throw new Error('Required columns missing: ' + result.missing.join(', ') + '. Expected: ' + REQUIRED_COLUMNS.join(', '));
        }
        if (!result.records.length) throw new Error('No data rows found in ' + file.name + '.');
        applyRecords(result.records, { name: file.name, syncedAt: Date.now() });
        toast('Imported ' + formatNum(result.records.length) + ' rows from ' + file.name, 'success');
      } catch (err) {
        showError(err.message || 'Failed to parse the uploaded file.');
      } finally {
        if (title && originalTitle) title.textContent = originalTitle;
        var input = $('fileInput');
        if (input) input.value = '';
      }
    };
    reader.onerror = function () {
      showError('Could not read the selected file.');
      if (title && originalTitle) title.textContent = originalTitle;
      var input = $('fileInput');
      if (input) input.value = '';
    };
    reader.readAsArrayBuffer(file);
  }

  function setSyncing(on) {
    var ids = ['syncBtn', 'syncBtn2'];
    for (var i = 0; i < ids.length; i++) {
      var btn = $(ids[i]);
      if (!btn) continue;
      var icon = btn.querySelector('svg, i');
      if (on) {
        btn.disabled = true;
        if (icon) icon.classList.add('spin');
      } else {
        btn.disabled = false;
        if (icon) icon.classList.remove('spin');
      }
    }
    setText('syncBtn2Label', on ? 'Syncing…' : 'Pull live from Google Sheet');
  }

  async function syncFromSheet() {
    if (syncing) return;
    syncing = true;
    setSyncing(true);
    try {
      var records;
      var via = 'csv';
      try {
        var res = await fetch(SHEET_CSV_URL);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        var csv = await res.text();
        var wb = XLSX.read(csv, { type: 'string' });
        var result = ingestWorkbook(wb);
        if (result.missing.length) throw new Error('Sheet columns not recognised: ' + result.missing.join(', '));
        if (!result.records.length) throw new Error('Sheet contains no data rows.');
        records = result.records;
      } catch (csvErr) {
        via = 'gviz';
        var gRes = await fetch(SHEET_GVIZ_URL);
        if (!gRes.ok) throw new Error('HTTP ' + gRes.status);
        var raw = parseGvizRows(await gRes.text());
        records = raw.map(normalizeRecord);
      }
      if (!records.length) throw new Error('Sheet contains no usable rows.');
      applyRecords(records, { name: 'Live Google Sheet', syncedAt: Date.now(), via: via });
      toast('Synced ' + formatNum(records.length) + ' rows from Google Sheet', 'success');
    } catch (err) {
      if (isFileProtocol()) showLocalNotice();
      showError('Google Sheet sync failed — ' + describeSyncError(err));
    } finally {
      setSyncing(false);
      syncing = false;
    }
  }

  /* -------------------------------------------------------------
     Events + init
     ------------------------------------------------------------- */
  function wireEvents() {
    var dz = $('dropzone');
    var fileInput = $('fileInput');
    var pickBtn = $('pickBtn');

    if (dz && fileInput) {
      dz.addEventListener('click', function (e) {
        if (e.target.closest && e.target.closest('button')) return;
        fileInput.click();
      });
      dz.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          fileInput.click();
        }
      });
      ['dragenter', 'dragover'].forEach(function (ev) {
        dz.addEventListener(ev, function (e) {
          e.preventDefault();
          e.stopPropagation();
          dz.classList.add('is-dragover');
        });
      });
      ['dragleave', 'drop'].forEach(function (ev) {
        dz.addEventListener(ev, function (e) {
          e.preventDefault();
          e.stopPropagation();
          dz.classList.remove('is-dragover');
        });
      });
      dz.addEventListener('drop', function (e) {
        var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
        if (file) handleFile(file);
      });
    }

    if (pickBtn && fileInput) {
      pickBtn.addEventListener('click', function () { fileInput.click(); });
    }
    if (fileInput) {
      fileInput.addEventListener('change', function () {
        if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
      });
    }

    var syncBtn = $('syncBtn');
    var syncBtn2 = $('syncBtn2');
    if (syncBtn) syncBtn.addEventListener('click', syncFromSheet);
    if (syncBtn2) syncBtn2.addEventListener('click', syncFromSheet);

    var priorityPills = document.querySelectorAll('#priorityFilter .pill');
    for (var i = 0; i < priorityPills.length; i++) {
      priorityPills[i].addEventListener('click', function (ev) {
        setPriorityFilter(ev.currentTarget.getAttribute('data-priority'));
      });
    }
    var ackPills = document.querySelectorAll('#ackFilter .pill');
    for (var j = 0; j < ackPills.length; j++) {
      ackPills[j].addEventListener('click', function (ev) {
        setAckFilter(ev.currentTarget.getAttribute('data-ack'));
      });
    }
    var resetBtn = $('resetFilters');
    if (resetBtn) resetBtn.addEventListener('click', function () { resetFilters(false); });

    var search = $('searchInput');
    if (search) {
      var debounce = null;
      search.addEventListener('input', function () {
        clearTimeout(debounce);
        debounce = setTimeout(function () {
          state.query = search.value.trim();
          renderAll();
        }, 140);
      });
    }

    var sortHeaders = document.querySelectorAll('th.sortable');
    for (var k = 0; k < sortHeaders.length; k++) {
      sortHeaders[k].addEventListener('click', function (ev) {
        var key = ev.currentTarget.getAttribute('data-key');
        if (state.sortKey === key) {
          state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sortKey = key;
          state.sortDir = (key === 'orderDate' || key === 'dispatchDate') ? 'desc' : 'asc';
        }
        renderTable();
        renderSortIcons();
      });
    }
  }

  async function init() {
    refreshIcons();
    wireEvents();
    createCharts();
    renderSortIcons();
    if (isFileProtocol()) showLocalNotice();

    try {
      db = await openDB();
    } catch (e) {
      db = null;
      console.warn('IndexedDB unavailable:', e);
    }

    var cached = await loadPersisted();
    if (cached && cached.records && cached.records.length) {
      state.records = cached.records;
      state.source = cached.meta && cached.meta.name ? { name: cached.meta.name, syncedAt: cached.meta.syncedAt } : null;
      renderAll();
      toast('Restored ' + formatNum(state.records.length) + ' rows from local cache', 'info');
    } else {
      renderAll();
    }
  }

  /* Expose pure core for debugging + tests */
  window.DashboardCore = {
    REQUIRED_COLUMNS: REQUIRED_COLUMNS,
    HEADER_ALIASES: HEADER_ALIASES,
    PLACE_TYPO_MAP: PLACE_TYPO_MAP,
    cleanPlace: cleanPlace,
    cleanCustomer: cleanCustomer,
    normalizePriority: normalizePriority,
    normalizeAck: normalizeAck,
    parseDate: parseDate,
    normalizeRecord: normalizeRecord,
    detectHeaders: detectHeaders,
    buildRecords: buildRecords,
    ingestWorkbook: ingestWorkbook,
    parseGvizRows: parseGvizRows,
    computeKpis: computeKpis,
    buildBarData: buildBarData,
    buildDoughnutData: buildDoughnutData,
    sortRows: sortRows,
    filteredRows: filteredRows
  };

  init();
})();
