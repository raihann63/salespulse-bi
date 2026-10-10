// ===== Config =====
const DATA_SHEET = 'business_sales_data';
const SOLD = ['complete', 'received'];     

// ===== Main =====
function buildSummary() {
  const t0 = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  Logger.log('Loaded in ' + (new Date() - t0) / 1000 + ' sec | data rows: ' + (rows.length - 1));

  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);
  Logger.log('Sample date: ' + rows[1][idx.order_date] + ' | sample total: ' + rows[1][idx.total]);

  const byState = {}, byCat = {}, byMonth = {}, bySku = {}, byCity = {}, stateCat = {};
  const regionOf = {};
  const allCusts = new Set();
  let soldRows = 0, soldRev = 0, badRows = 0;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (SOLD.indexOf(r[idx.status]) === -1) continue;

    const rev = parseFloat(String(r[idx.total]).replace(/[^0-9.\-]/g, ''));
    if (isNaN(rev)) { badRows++; continue; }

    const state = r[idx.State], cat = r[idx.category], sku = r[idx.sku];
    const city = r[idx.City] + ', ' + state, cust = r[idx.cust_id];

    soldRows++; soldRev += rev; allCusts.add(cust);
    regionOf[state] = r[idx.Region];
    add_(byState, state, rev, cust);
    add_(byCat, cat, rev, cust);
    add_(byMonth, monthKey_(r[idx.order_date]), rev, cust);
    add_(bySku, sku, rev, null);
    add_(byCity, city, rev, cust);
    if (!stateCat[state]) stateCat[state] = {};
    stateCat[state][cat] = (stateCat[state][cat] || 0) + rev;
  }

  Logger.log('Sold items: ' + soldRows + ' | Sold revenue: ' + soldRev.toFixed(2) +
             ' | Unique customers: ' + allCusts.size + ' | Skipped bad rows: ' + badRows);

  const H = (first, hasCusts) =>
    hasCusts ? [first, 'revenue', 'items_sold', 'unique_customers'] : [first, 'revenue', 'items_sold'];

  const stateRows = toRows_(byState, true).map(r => [r[0], regionOf[r[0]], r[1], r[2], r[3]]);
  writeTable_(ss, 'sum_state', ['state', 'region', 'revenue', 'items_sold', 'unique_customers'], stateRows);
  writeTable_(ss, 'sum_category', H('category', true), toRows_(byCat, true));
  writeTable_(ss, 'sum_month', H('month', true),
              toRows_(byMonth, true).sort((a, b) => a[0] < b[0] ? -1 : 1));
  writeTable_(ss, 'sum_sku_top50', H('sku', false), toRows_(bySku, false).slice(0, 50));
  const MIN_CITY_CUST = 10;
  const cityRows = toRows_(byCity, true).filter(r => r[3] >= MIN_CITY_CUST).slice(0, 50);
  writeTable_(ss, 'sum_city_top50', H('city', true), cityRows);

  const cats = Object.keys(byCat).sort();
  const matrix = Object.keys(stateCat).sort()
    .map(s => [s].concat(cats.map(c => round_(stateCat[s][c] || 0))));
  writeTable_(ss, 'sum_state_category', ['state'].concat(cats), matrix);

  writeTable_(ss, 'sum_kpi', ['metric', 'value'], [
    ['sold_statuses', SOLD.join(' + ')],
    ['sold_items', soldRows],
    ['sold_revenue', round_(soldRev)],
    ['unique_customers_sold', allCusts.size],
    ['rows_skipped_bad_total', badRows],
    ['updated_at', new Date().toISOString()]
  ]);

    CacheService.getScriptCache().remove('dash_v1');   // dashboard-এর পুরোনো cache মুছে ফেলে
  Logger.log('Done in ' + (new Date() - t0) / 1000 + ' sec');

}

// ===== Helpers =====
function loadCsv_(ss, sheet) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() +
              '/export?format=csv&gid=' + sheet.getSheetId();
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('CSV export failed: HTTP ' + res.getResponseCode());
  return Utilities.parseCsv(res.getContentText());
}

function add_(map, key, rev, cust) {
  let o = map[key];
  if (!o) o = map[key] = { rev: 0, rows: 0, custs: cust !== null ? new Set() : null };
  o.rev += rev;
  o.rows++;
  if (cust !== null) o.custs.add(cust);
}

function toRows_(map, hasCusts) {
  return Object.keys(map).map(k => {
    const o = map[k];
    const row = [k, round_(o.rev), o.rows];
    if (hasCusts) row.push(o.custs.size);
    return row;
  }).sort((a, b) => b[1] - a[1]);
}

function monthKey_(s) {
  s = String(s).trim();
  let m = /^(\d{4})-(\d{1,2})/.exec(s);        // 2020-10-01
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = /^(\d{1,2})\/\d{1,2}\/(\d{4})/.exec(s);  // 10/1/2020 (মাস/দিন/বছর)
  if (m) return m[2] + '-' + ('0' + m[1]).slice(-2);
  return 'unknown';
}

function round_(x) { return Math.round(x * 100) / 100; }

function writeTable_(ss, name, header, rows) {
  let sh = ss.getSheetByName(name);
  if (sh) sh.clear(); else sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
}




// ===== Web App =====
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('SalesPulse BI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getCube() {
  return loadCube_();
}

function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('dash_v1');
  if (hit) return JSON.parse(hit);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const read = (name, fmt) => clean_(ss.getSheetByName(name).getDataRange().getValues(), tz, fmt);

  const data = {
    kpi: read('sum_kpi', 'yyyy-MM-dd HH:mm'),
    state: read('sum_state', 'yyyy-MM'),
    category: read('sum_category', 'yyyy-MM'),
    month: read('sum_month', 'yyyy-MM'),
    sku: read('sum_sku_top50', 'yyyy-MM'),
    city: read('sum_city_top50', 'yyyy-MM'),
    matrix: read('sum_state_category', 'yyyy-MM')
  };
  try { cache.put('dash_v1', JSON.stringify(data), 21600); } catch (e) {}
  return data;
}

// Sheets কখনো "2020-10" বা ISO লেখাকে Date বানিয়ে ফেলে; Date ব্রাউজারে পাঠানো যায় না
function clean_(rows, tz, fmt) {
  return rows.map(r => r.map(v => v instanceof Date ? Utilities.formatDate(v, tz, fmt) : v));
}



function checkMonths() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);
  const m = {};
  for (let i = 1; i < rows.length; i++) {
    const k = monthKey_(rows[i][idx.order_date]);
    if (!m[k]) m[k] = { total: 0, complete: 0, canceled: 0 };
    m[k].total++;
    if (rows[i][idx.status] === 'complete') m[k].complete++;
    if (rows[i][idx.status] === 'canceled') m[k].canceled++;
  }
  Object.keys(m).sort().forEach(k =>
    Logger.log(k + ' | total: ' + m[k].total + ' | complete: ' + m[k].complete + ' | canceled: ' + m[k].canceled));
}


function checkStatusByMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);
  const m = {};
  for (let i = 1; i < rows.length; i++) {
    const k = monthKey_(rows[i][idx.order_date]);
    const s = rows[i][idx.status];
    if (!m[k]) m[k] = {};
    m[k][s] = (m[k][s] || 0) + 1;
  }
  Object.keys(m).sort().forEach(k => {
    const top = Object.entries(m[k]).sort((a, b) => b[1] - a[1]).slice(0, 5);
    Logger.log(k + ' | ' + JSON.stringify(top));
  });
}


function checkOutliers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  const sold = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (SOLD.indexOf(r[idx.status]) === -1) continue;
    sold.push({
      total: parseFloat(r[idx.total]) || 0,
      sku: r[idx.sku],
      city: r[idx.City] + ', ' + r[idx.State],
      cust: r[idx.cust_id],
      price: r[idx.price]
    });
  }

  sold.sort((a, b) => b.total - a.total);
  sold.slice(0, 15).forEach(o =>
    Logger.log(o.total + ' | ' + o.sku + ' | ' + o.city + ' | cust ' + o.cust + ' | unit price ' + o.price));

  const noxa = sold.filter(o => o.city === 'Noxapater, MS');
  Logger.log('Noxapater rows: ' + noxa.length + ' | customers: ' + new Set(noxa.map(o => o.cust)).size);
}


// ===== Step 1: Data profiling =====
function profileData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) throw new Error('Tab পাওয়া যায়নি: ' + DATA_SHEET);

  const rows = loadCsv_(ss, sheet);
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  let blankOrder = 0, blankCust = 0, badDate = 0, badTotal = 0, negTotal = 0, zeroTotal = 0, badPrice = 0, dupLines = 0;
  let minDate = '9999', maxDate = '0000';
  const seen = new Set(), orders = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const oid = r[idx.order_id], d = r[idx.order_date];
    const total = parseFloat(r[idx.total]), price = parseFloat(r[idx.price]);

    if (!oid) blankOrder++;
    if (!r[idx.cust_id]) blankCust++;
    if (!/^\d{4}-\d{2}-\d{2}/.test(d)) badDate++;
    else { if (d < minDate) minDate = d; if (d > maxDate) maxDate = d; }
    if (isNaN(total)) badTotal++; else { if (total < 0) negTotal++; if (total === 0) zeroTotal++; }
    if (isNaN(price)) badPrice++;

    const key = [oid, r[idx.sku], r[idx.total], r[idx.status], r[idx.cust_id]].join('|');
    if (seen.has(key)) dupLines++; else seen.add(key);

    if (!orders[oid]) orders[oid] = [];
    orders[oid].push({ sku: r[idx.sku], price: price, total: total, status: r[idx.status] });
  }

  const ids = Object.keys(orders);
  let single = 0, multi = 0, sameTotals = 0, diffTotals = 0, mixedStatus = 0;
  const samples = [];
  ids.forEach(id => {
    const L = orders[id];
    if (new Set(L.map(x => x.status)).size > 1) mixedStatus++;
    if (L.length === 1) { single++; return; }
    multi++;
    if (new Set(L.map(x => x.total)).size === 1) sameTotals++; else diffTotals++;
    if (samples.length < 3) samples.push(id + ': ' + L.map(x => x.sku + ' price ' + x.price + ' total ' + x.total).join(' || '));
  });

  Logger.log('Headers: ' + rows[0].join(', '));
  Logger.log('Data rows: ' + (rows.length - 1) + ' | distinct order_id: ' + ids.length);
  Logger.log('Blank order_id: ' + blankOrder + ' | blank cust_id: ' + blankCust);
  Logger.log('Bad date: ' + badDate + ' | date range: ' + minDate + ' to ' + maxDate);
  Logger.log('Bad total: ' + badTotal + ' | negative: ' + negTotal + ' | zero: ' + zeroTotal + ' | bad price: ' + badPrice);
  Logger.log('Possible duplicate lines: ' + dupLines);
  Logger.log('Orders with 1 line: ' + single + ' | multi-line: ' + multi);
  Logger.log('Multi-line orders, same total on every line: ' + sameTotals + ' | different totals: ' + diffTotals);
  Logger.log('Orders with mixed statuses across lines: ' + mixedStatus);
  samples.forEach(s => Logger.log('Sample ' + s));
}

function profileData2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  const zeroByStatus = {}, orderStatus = {}, ordersByStatus = {}, orders = {};
  let intUnits = 0, fracUnits = 0;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const oid = r[idx.order_id], st = r[idx.status];
    const total = parseFloat(r[idx.total]), price = parseFloat(r[idx.price]);

    if (total === 0) zeroByStatus[st] = (zeroByStatus[st] || 0) + 1;
    orderStatus[oid] = st;

    if (total > 0 && price > 0) {
      const u = total / price;
      if (Math.abs(u - Math.round(u)) < 0.01) intUnits++; else fracUnits++;
    }

    if (!orders[oid]) orders[oid] = [];
    orders[oid].push({ sku: r[idx.sku], price: price, total: total });
  }

  Object.keys(orderStatus).forEach(id => {
    const s = orderStatus[id];
    ordersByStatus[s] = (ordersByStatus[s] || 0) + 1;
  });

  let sameTotal = 0, zeroOnly = 0, suspicious = 0;
  const samples = [];
  Object.keys(orders).forEach(id => {
    const L = orders[id];
    if (L.length < 2) return;
    if (new Set(L.map(x => x.total)).size !== 1) return;
    if (L[0].total === 0) { zeroOnly++; return; }
    sameTotal++;
    if (new Set(L.map(x => x.price)).size > 1) {   // দাম আলাদা কিন্তু total এক: সন্দেহজনক
      suspicious++;
      if (samples.length < 5)
        samples.push(id + ': ' + L.map(x => x.sku + ' price ' + x.price + ' total ' + x.total).join(' || '));
    }
  });

  Logger.log('Zero-total rows by status: ' + JSON.stringify(zeroByStatus));
  Logger.log('Distinct orders by status: ' + JSON.stringify(ordersByStatus));
  Logger.log('total/price whole number: ' + intUnits + ' | fractional: ' + fracUnits);
  Logger.log('Same-total multi-line orders (non-zero): ' + sameTotal + ' | all-zero: ' + zeroOnly +
             ' | different prices but same total: ' + suspicious);
  samples.forEach(s => Logger.log('Sample ' + s));
}

// ===== Step 3: Business rules =====
const STATUS_GROUPS = {
  completed:    ['complete', 'received'],
  cancelled:    ['canceled'],
  refunded:     ['order_refunded', 'refund', 'closed'],
  in_progress:  ['pending', 'pending_paypal', 'processing', 'holded', 'payment_review'],
  unclassified: ['paid', 'cod']
};

function groupOf_(status) {
  for (const g in STATUS_GROUPS) if (STATUS_GROUPS[g].indexOf(status) !== -1) return g;
  return 'unknown';
}

function buildStep3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  const orderGroup = {}, orderRev = {}, custOrders = {};
  const value = { completed: 0, cancelled: 0, refunded: 0 };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const oid = r[idx.order_id], g = groupOf_(r[idx.status]);
    const total = parseFloat(r[idx.total]) || 0;
    orderGroup[oid] = g;
    if (value[g] !== undefined) value[g] += total;
    if (g === 'completed') {
      orderRev[oid] = (orderRev[oid] || 0) + total;
      const c = r[idx.cust_id];
      if (!custOrders[c]) custOrders[c] = new Set();
      custOrders[c].add(oid);
    }
  }

  const n = { completed: 0, cancelled: 0, refunded: 0, in_progress: 0, unclassified: 0, unknown: 0 };
  Object.keys(orderGroup).forEach(id => n[orderGroup[id]]++);
  const totalOrders = Object.keys(orderGroup).length;
  const eligible = n.completed + n.cancelled + n.refunded;
  const zeroRevOrders = Object.keys(orderRev).filter(id => orderRev[id] === 0).length;
  const custIds = Object.keys(custOrders);
  const repeat = custIds.filter(c => custOrders[c].size >= 2).length;
  const pct = (a, b) => b ? round_(a / b * 100) : 'N/A';

  writeTable_(ss, 'KPI_Summary', ['metric', 'value', 'definition'], [
    ['total_orders', totalOrders, 'Distinct order_id'],
    ['completed_orders', n.completed, 'Status: complete + received'],
    ['cancelled_orders', n.cancelled, 'Status: canceled'],
    ['refunded_orders', n.refunded, 'Status: order_refunded + refund + closed'],
    ['in_progress_orders', n.in_progress, 'pending, pending_paypal, processing, holded, payment_review (excluded from rates)'],
    ['unclassified_orders', n.unclassified, 'paid, cod: payment words inside status column (excluded from rates)'],
    ['eligible_orders', eligible, 'completed + cancelled + refunded'],
    ['cancellation_rate_pct', pct(n.cancelled, eligible), 'cancelled / eligible x 100'],
    ['refund_rate_pct', pct(n.refunded, eligible), 'refunded / eligible x 100'],
    ['revenue_completed', round_(value.completed), 'Sum of total over completed orders'],
    ['cancelled_value', round_(value.cancelled), 'Sum of total over cancelled orders (lost sales)'],
    ['aov', n.completed ? round_(value.completed / n.completed) : 'N/A', 'revenue_completed / completed_orders'],
    ['completed_orders_zero_revenue', zeroRevOrders, 'Completed orders where every line has total = 0'],
    ['active_customers', custIds.length, 'Customers with at least one completed order'],
    ['repeat_customer_rate_pct', pct(repeat, custIds.length), 'Customers with 2+ distinct completed orders / active customers'],
    ['updated_at', new Date().toISOString(), '']
  ]);

  writeTable_(ss, 'Data_Dictionary', ['column', 'meaning', 'type', 'rule'], [
    ['order_id', 'Order identifier. One order can have several rows (lines)', 'text', 'Count DISTINCT for orders'],
    ['order_date', 'Order date, 2020-10-01 to 2021-09-30 (validated, no bad dates)', 'date text yyyy-MM-dd', 'Group by month'],
    ['status', 'Raw order status, 13 values. All lines of an order share one status (validated)', 'text', 'Mapped to groups by STATUS_GROUPS'],
    ['sku', 'Product code. Size/variant is often part of the code', 'text', 'Variants count as separate SKUs'],
    ['price', 'Unit price of the line', 'number', 'Not used for revenue'],
    ['total', 'Line-level sales value (validated). About 6.7% of rows are 0', 'number', 'Revenue = sum of total over completed orders'],
    ['category', 'Product category', 'text', 'Category analysis'],
    ['payment_method', 'Payment method', 'text', 'Payment analysis'],
    ['cust_id', 'Customer identifier', 'text', 'Count DISTINCT for customers'],
    ['Gender', 'M or F (validated)', 'text', 'Customer segmentation'],
    ['age', 'Customer age, 18 to 75 (validated)', 'number', 'Age groups'],
    ['Place Name, County, City', 'Finer location fields', 'text', 'Use City together with State'],
    ['State, Region', 'US state code (51 values incl. DC) and region', 'text', 'Geographic analysis']
  ]);

  Logger.log('Orders: ' + totalOrders + ' | ' + JSON.stringify(n));
  Logger.log('Eligible: ' + eligible + ' | Cancellation %: ' + pct(n.cancelled, eligible) + ' | Refund %: ' + pct(n.refunded, eligible));
  Logger.log('Revenue: ' + round_(value.completed) + ' | Cancelled value: ' + round_(value.cancelled) + ' | AOV: ' + round_(value.completed / n.completed));
  Logger.log('Completed zero-revenue orders: ' + zeroRevOrders + ' | Active customers: ' + custIds.length + ' | Repeat: ' + repeat);
}


function ageGroup_(a) {
  a = parseInt(a, 10);
  if (isNaN(a)) return 'unknown';
  if (a <= 25) return '18-25';
  if (a <= 35) return '26-35';
  if (a <= 45) return '36-45';
  if (a <= 60) return '46-60';
  return '61+';
}

function cubeCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  const first = {}, conflicts = { month: 0, state: 0, payment: 0, customer: 0 };
  const cats = {}, pays = {}, A = new Set(), B = new Set(), C = new Set();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i], oid = r[idx.order_id];
    const cur = { m: monthKey_(r[idx.order_date]), st: r[idx.State], pm: r[idx.payment_method], cu: r[idx.cust_id] };
    let f = first[oid];
    if (!f) { f = first[oid] = cur; pays[cur.pm] = (pays[cur.pm] || 0) + 1; }
    else {
      if (f.m !== cur.m) conflicts.month++;
      if (f.st !== cur.st) conflicts.state++;
      if (f.pm !== cur.pm) conflicts.payment++;
      if (f.cu !== cur.cu) conflicts.customer++;
    }
    const g = groupOf_(r[idx.status]), cat = r[idx.category];
    cats[cat] = (cats[cat] || 0) + 1;
    const base = [f.m, f.st, f.pm, g];
    C.add(base.join('|'));
    const a = base.concat(cat).join('|');
    A.add(a);
    B.add(a + '|' + r[idx.Gender] + '|' + ageGroup_(r[idx.age]));
  }

  Logger.log('Conflicts inside orders (should be 0): ' + JSON.stringify(conflicts));
  Logger.log('Orders by payment method: ' + JSON.stringify(pays));
  Logger.log('Categories (lines): ' + JSON.stringify(cats));
  Logger.log('Cube cells | order-level (month,state,payment,status): ' + C.size +
             ' | + category: ' + A.size + ' | + category,gender,age group: ' + B.size);
}


// ===== Step 4a: Summary cubes =====
function ageGroup_(a) {
  a = parseInt(a, 10);
  if (isNaN(a)) return 'unknown';
  if (a <= 25) return '18-25';
  if (a <= 35) return '26-35';
  if (a <= 45) return '36-45';
  if (a <= 60) return '46-60';
  return '61+';
}

function buildCubes() {
  const t0 = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  Logger.log('Loaded in ' + (new Date() - t0) / 1000 + ' sec');
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  const D = { month: [], state: [], pay: [], cat: [], gender: [], age: [],
              group: ['completed', 'cancelled', 'refunded', 'in_progress', 'unclassified', 'unknown'] };
  const M = { month: {}, state: {}, pay: {}, cat: {}, gender: {}, age: {}, group: {} };
  D.group.forEach((g, i) => M.group[g] = i);
  const ix = (k, v) => { if (!(v in M[k])) { M[k][v] = D[k].length; D[k].push(v); } return M[k][v]; };

  const ord = {}, cat = {}, cus = {}, first = {}, seenOC = new Set(), regionOf = {};
  let genderConf = 0, ageConf = 0;
  const bump = (map, key, o, l, rev) => {
    const x = map[key] || (map[key] = [0, 0, 0]);
    x[0] += o; x[1] += l; x[2] += rev;
  };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i], oid = r[idx.order_id];
    const total = parseFloat(r[idx.total]) || 0;
    let f = first[oid];
    const isNew = !f;
    if (isNew) {
      f = first[oid] = [
        ix('month', monthKey_(r[idx.order_date])), ix('state', r[idx.State]), ix('pay', r[idx.payment_method]),
        ix('group', groupOf_(r[idx.status])), ix('gender', r[idx.Gender]), ix('age', ageGroup_(r[idx.age]))
      ];
      regionOf[r[idx.State]] = r[idx.Region];
    } else {
      if (D.gender[f[4]] !== r[idx.Gender]) genderConf++;
      if (D.age[f[5]] !== ageGroup_(r[idx.age])) ageConf++;
    }
    const base = f[0] + ',' + f[1] + ',' + f[2] + ',' + f[3];
    bump(ord, base, isNew ? 1 : 0, 1, total);

    const c = ix('cat', r[idx.category]), oc = oid + '|' + c;
    const newOC = !seenOC.has(oc);
    if (newOC) seenOC.add(oc);
    bump(cat, base + ',' + c, newOC ? 1 : 0, 1, total);

    bump(cus, f[0] + ',' + f[3] + ',' + f[4] + ',' + f[5], isNew ? 1 : 0, 1, total);
  }
  Logger.log('Aggregated in ' + (new Date() - t0) / 1000 + ' sec');

  const toArr = map => Object.keys(map).map(k =>
    k.split(',').map(Number).concat([map[k][0], map[k][1], Math.round(map[k][2] * 10000) / 10000]));
  const cube = { dims: D, regionOf: regionOf, ord: toArr(ord), cat: toArr(cat), cus: toArr(cus),
                 built: new Date().toISOString() };

  // যাচাই: ord = [month,state,pay,group,orders,lines,rev]  cat = [month,state,pay,group,cat,orders,lines,rev]
  const byGroup = {};
  cube.ord.forEach(a => {
    const g = D.group[a[3]];
    const c = byGroup[g] || (byGroup[g] = { orders: 0, rev: 0 });
    c.orders += a[4]; c.rev += a[6];
  });
  let catCompletedRev = 0;
  cube.cat.forEach(a => { if (a[3] === 0) catCompletedRev += a[7]; });
  Object.keys(byGroup).forEach(g => byGroup[g].rev = round_(byGroup[g].rev));
  Logger.log('ord cube by group: ' + JSON.stringify(byGroup));
  Logger.log('cat cube completed revenue: ' + round_(catCompletedRev));
  Logger.log('Rows | ord: ' + cube.ord.length + ' | cat: ' + cube.cat.length + ' | cus: ' + cube.cus.length);
  Logger.log('Gender/age conflicts inside orders (should be 0): ' + genderConf + ' / ' + ageConf);

  // cube_json tab-এ ছোট ছোট অংশে লেখা (প্রতি অংশের শুরুতে 'J', যেন Sheets কিছু বদলে না ফেলে)
  const json = JSON.stringify(cube), CH = 40000, parts = [];
  for (let i = 0; i < json.length; i += CH) parts.push([parts.length + 1, 'J' + json.substring(i, i + CH)]);
  writeTable_(ss, 'cube_json', ['part', 'json'], parts);
  Logger.log('JSON chars: ' + json.length + ' | parts: ' + parts.length + ' | done in ' + (new Date() - t0) / 1000 + ' sec');
}

function loadCube_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('cube_json');
  const n = sh.getLastRow() - 1;
  const parts = sh.getRange(2, 2, n, 1).getValues().map(r => String(r[0]).substring(1));
  return JSON.parse(parts.join(''));
}

function testCubeRead() {
  const t = new Date();
  const cube = loadCube_();
  Logger.log('Cube read+parse: ' + (new Date() - t) / 1000 + ' sec | ord: ' + cube.ord.length +
             ' | cat: ' + cube.cat.length + ' | cus: ' + cube.cus.length + ' | built: ' + cube.built);
}



// =====================================================================
// Code_additions.gs  -  paste at the END of your existing Code.gs
// (nothing in the existing code is changed)
//
// Step 1: run buildExtras() once (and again whenever the raw data changes)
// Step 2: Deploy > Manage deployments > edit > New version
// =====================================================================

const EXTRA_SKU_TOP = 150;     // top SKUs by revenue + top SKUs by completed orders (union is kept)
const EXTRA_CITY_TOP = 400;    // cities with the most eligible orders

// ===== Extras: product x month x region and city x month tables =====
function buildExtras() {
  const t0 = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = loadCsv_(ss, ss.getSheetByName(DATA_SHEET));
  Logger.log('Loaded in ' + (new Date() - t0) / 1000 + ' sec | data rows: ' + (rows.length - 1));
  const idx = {};
  rows[0].forEach((h, i) => idx[h] = i);

  const first = {};                 // order_id -> shared order info (validated: one month/state/status per order)
  const seenPair = new Set();       // order_id|sku  (counts each order once per product)
  const prod = {}, skuCat = {}, skuTot = {};
  const city = {}, cityTot = {}, cityMeta = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i], oid = r[idx.order_id];
    const total = parseFloat(r[idx.total]) || 0;
    let f = first[oid];
    const isNew = !f;
    if (isNew) {
      const st = r[idx.State];
      f = first[oid] = { m: monthKey_(r[idx.order_date]), rg: r[idx.Region], st: st,
                         ci: r[idx.City] + ', ' + st, g: groupOf_(r[idx.status]) };
    }
    const g = f.g;

    // ---- product ----
    const sku = r[idx.sku];
    if (!(sku in skuCat)) skuCat[sku] = r[idx.category];
    const pair = oid + '|' + sku;
    const newPair = !seenPair.has(pair);
    if (newPair) seenPair.add(pair);
    // [completedOrders, completedLines, revenue, cancelledOrders, refundedOrders]
    const p = prod[sku + '\t' + f.m + '\t' + f.rg] || (prod[sku + '\t' + f.m + '\t' + f.rg] = [0, 0, 0, 0, 0]);
    if (g === 'completed') {
      if (newPair) p[0]++;
      p[1]++; p[2] += total;
      const t = skuTot[sku] || (skuTot[sku] = [0, 0]);
      t[0] += total; if (newPair) t[1]++;
    } else if (g === 'cancelled') { if (newPair) p[3]++; }
    else if (g === 'refunded') { if (newPair) p[4]++; }

    // ---- city (one count per order) ----
    // [completedOrders, cancelledOrders, refundedOrders, revenue]
    const ck = f.ci + '\t' + f.m;
    const c = city[ck] || (city[ck] = [0, 0, 0, 0]);
    if (isNew) {
      if (g === 'completed') c[0]++;
      else if (g === 'cancelled') c[1]++;
      else if (g === 'refunded') c[2]++;
      if (g === 'completed' || g === 'cancelled' || g === 'refunded') cityTot[f.ci] = (cityTot[f.ci] || 0) + 1;
    }
    if (g === 'completed') c[3] += total;
    if (!cityMeta[f.ci]) cityMeta[f.ci] = [f.st, f.rg];
  }
  Logger.log('Aggregated in ' + (new Date() - t0) / 1000 + ' sec');

  // ---- choose which SKUs / cities to keep ----
  const skuKeys = Object.keys(skuTot);
  const byRev = skuKeys.slice().sort((a, b) => skuTot[b][0] - skuTot[a][0]).slice(0, EXTRA_SKU_TOP);
  const byOrd = skuKeys.slice().sort((a, b) => skuTot[b][1] - skuTot[a][1]).slice(0, EXTRA_SKU_TOP);
  const keepSku = new Set(byRev.concat(byOrd));
  const skus = Array.from(keepSku), skuIx = {};
  skus.forEach((s, i) => skuIx[s] = i);

  const cities = Object.keys(cityTot).sort((a, b) => cityTot[b] - cityTot[a]).slice(0, EXTRA_CITY_TOP);
  const cityIx = {};
  cities.forEach((c, i) => cityIx[c] = i);

  const months = [], mIx = {}, regions = [], rIx = {};
  const ix = (arr, map, v) => { if (!(v in map)) { map[v] = arr.length; arr.push(v); } return map[v]; };

  const prodOut = [];
  Object.keys(prod).forEach(k => {
    const parts = k.split('\t');
    if (!keepSku.has(parts[0])) return;
    const p = prod[k];
    prodOut.push([skuIx[parts[0]], ix(months, mIx, parts[1]), ix(regions, rIx, parts[2]),
                  p[0], p[1], round_(p[2]), p[3], p[4]]);
  });

  const cityOut = [];
  Object.keys(city).forEach(k => {
    const parts = k.split('\t');
    if (!(parts[0] in cityIx)) return;
    const x = city[k];
    cityOut.push([cityIx[parts[0]], ix(months, mIx, parts[1]), x[0], x[1], x[2], round_(x[3])]);
  });

  const out = {
    built: new Date().toISOString(),
    months: months, regions: regions,
    skus: skus, skuCat: skus.map(s => skuCat[s]),
    cities: cities, cityState: cities.map(c => cityMeta[c][0]), cityRegion: cities.map(c => cityMeta[c][1]),
    prod: prodOut,      // [sku, month, region, completedOrders, completedLines, revenue, cancelledOrders, refundedOrders]
    city: cityOut       // [city, month, completedOrders, cancelledOrders, refundedOrders, revenue]
  };

  const json = JSON.stringify(out), CH = 40000, parts = [];
  for (let i = 0; i < json.length; i += CH) parts.push([parts.length + 1, 'J' + json.substring(i, i + CH)]);
  writeTable_(ss, 'extra_json', ['part', 'json'], parts);
  Logger.log('SKUs kept: ' + skus.length + ' | cities kept: ' + cities.length +
             ' | prod rows: ' + prodOut.length + ' | city rows: ' + cityOut.length);
  Logger.log('JSON chars: ' + json.length + ' | parts: ' + parts.length + ' | done in ' + (new Date() - t0) / 1000 + ' sec');
}

function loadExtras_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('extra_json');
  if (!sh || sh.getLastRow() < 2) throw new Error('extra_json tab not found. Run buildExtras() once first.');
  const n = sh.getLastRow() - 1;
  const parts = sh.getRange(2, 2, n, 1).getValues().map(r => String(r[0]).substring(1));
  return JSON.parse(parts.join(''));
}

function getExtras() {
  return loadExtras_();
}

function testExtrasRead() {
  const t = new Date();
  const x = loadExtras_();
  Logger.log('Extras read+parse: ' + (new Date() - t) / 1000 + ' sec | prod: ' + x.prod.length +
             ' | city: ' + x.city.length + ' | built: ' + x.built);
}
