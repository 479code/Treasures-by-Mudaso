/*
 * Treasures by Mudaso — business module (v70)
 * Cost prices, profit, profit progress, best categories, expenses and stock value.
 * Loads after app.js and extends it.
 */
const expensesStorageKey = 'treasures-expenses';
const stockBaseStorageKey = 'treasures-stock-base';
const profitTargetKey = 'treasures-profit-target';
const biPeriodKey = 'treasures-bi-period';
const expenseCategories = ['Transport', 'Packaging', 'Rent', 'Salaries', 'Utilities', 'Marketing', 'Bank charges', 'Other'];
const biPeriods = [['week', 'This week'], ['month', 'This month'], ['30d', '30 days'], ['year', 'This year'], ['all', 'All time']];
let expenses = readJson(expensesStorageKey, []);
let stockBases = readJson(stockBaseStorageKey, {});
let biPeriod = localStorage.getItem(biPeriodKey) || 'month';
let ownerSyncedThisSession = false;

function readJson(key, fallback) { try { return JSON.parse(localStorage.getItem(key) || 'null') ?? fallback; } catch (e) { return fallback; } }
function parseJson(value, fallback) { try { return JSON.parse(value || 'null') ?? fallback; } catch (e) { return fallback; } }
function esc(value) { return String(value ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function saveExpenses() { localStorage.setItem(expensesStorageKey, JSON.stringify(expenses)); }
function saveStockBases() { localStorage.setItem(stockBaseStorageKey, JSON.stringify(stockBases)); }
function dayKey(value) { const d = new Date(value); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }
function isoOrEmpty(value) { if (!value) return ''; const d = new Date(value); return isNaN(d) ? '' : d.toISOString(); }
function pct(part, whole) { return whole ? Math.round(part / whole * 100) : 0; }
function shortMoney(value) { const n = Number(value) || 0; const a = Math.abs(n); const s = n < 0 ? '−' : ''; if (a >= 1e6) return `${s}₦${(a / 1e6).toFixed(a >= 1e7 ? 0 : 1)}m`; if (a >= 1e3) return `${s}₦${Math.round(a / 1e3)}k`; return `${s}₦${a}`; }

/* ---------- Cost prices on products and packages (owner only, never shown to customers) ---------- */

const productRecordBase = productRecord;
productRecord = function (item) {
  const record = productRecordBase(item);
  record['Cost price (₦)'] = Number(item.costPrice || 0);
  if (stockBases[item.id] !== undefined) record.__stockBase = stockBases[item.id];
  stockBases[item.id] = Number(item.stock || 0);
  saveStockBases();
  return record;
};

const packageRecordBase = packageRecord;
packageRecord = function (item) { return { ...packageRecordBase(item), 'Cost price (₦)': Number(item.costPrice || 0) }; };

const addProductBase = addProductV2;
addProductV2 = function (index) {
  const item = Number.isInteger(index) ? products[index] : null;
  return addProductBase(index).replace('<label class="toggle"><input name="onSale"', `<label>Cost price (₦) <span>(what you pay per unit — customers never see this)</span><input name="costPrice" type="number" min="0" value="${Number(item?.costPrice) || ''}" placeholder="0"/></label><label class="toggle"><input name="onSale"`);
};

const createPackageBase = createPackage;
createPackage = function (index) {
  const item = Number.isInteger(index) ? featuredPackages[index] : null;
  return createPackageBase(index).replace('<label>Short label', `<label>Cost price (₦) <span>(your total cost to put one package together)</span><input name="costPrice" type="number" min="0" value="${Number(item?.costPrice) || ''}" placeholder="0"/></label><label>Short label`);
};

/* ---------- Owner data from the Sheet ---------- */

function onOwnerData(data) {
  const byId = (rows, key) => new Map((rows || []).map(row => [String(row[key]), row]));
  const productRows = byId(data.products, 'Product ID');
  products.forEach(item => {
    const row = productRows.get(String(item.id));
    if (!row) return;
    item.costPrice = Number(row['Cost price (₦)'] || 0);
    item.stock = Number(row.Stock || 0);
    stockBases[item.id] = item.stock;
  });
  saveProducts(); saveStockBases();
  const packageRows = byId(data.packages, 'Package ID');
  featuredPackages.forEach(item => {
    const row = packageRows.get(String(item.id));
    if (!row) return;
    item.costPrice = Number(row['Cost price (₦)'] || 0);
    item.available = String(row.Available).toLowerCase() !== 'false';
    item.installments = Number(row['Pay small small payments'] || 0);
  });
  savePackages();
  const orderRows = byId(data.orders, 'Order ID');
  orders.forEach(item => {
    const row = orderRows.get(String(item.id));
    if (!row) return;
    item.subtotal = Number(row['Subtotal (₦)'] || 0);
    item.cost = row['Cost (₦)'] === '' || row['Cost (₦)'] == null ? null : Number(row['Cost (₦)']);
    item.lines = parseJson(row.Lines, []);
    item.createdAt = isoOrEmpty(row['Created at']) || item.createdAt;
    item.paidAt = isoOrEmpty(row['Paid at']) || (item.status !== 'Awaiting payment' ? item.createdAt : '');
  });
  saveOrders();
  if (Array.isArray(data.expenses)) {
    expenses = data.expenses.filter(row => row['Expense ID'] && String(row.Deleted).toLowerCase() !== 'true').map(row => ({ id: String(row['Expense ID']), date: row.Date ? dayKey(row.Date) : dayKey(Date.now()), category: row.Category || 'Other', description: String(row.Description || ''), amount: Number(row['Amount (₦)'] || 0) }));
    saveExpenses();
  }
}

/* ---------- Money maths ---------- */

const isPaidOrder = order => order.status !== 'Awaiting payment';
const orderDate = order => new Date(order.paidAt || order.createdAt || 0);
const expenseDate = expense => new Date(`${expense.date}T12:00:00`);
const inRange = (date, start, end) => (!start || date >= start) && date <= end;

function catalogueCost(item) {
  if (!item) return 0;
  if (Number(item.costPrice)) return Number(item.costPrice);
  return (item.components || []).reduce((sum, c) => sum + Number(c.quantity || 0) * Number(products.find(p => p.id === c.productId)?.costPrice || 0), 0);
}

function orderLines(order) {
  if (Array.isArray(order.lines) && order.lines.length) return order.lines.map(line => ({ ...line, qty: Number(line.qty || 0), price: Number(line.price || 0), cost: Number(line.cost || 0) }));
  return String(order.items || '').split(', ').filter(Boolean).map(part => {
    const match = part.match(/^(\d+)\s*×\s*(.+)$/);
    const qty = match ? Number(match[1]) : 1;
    const name = (match ? match[2] : part).trim();
    const product = products.find(p => p.name === name);
    const pack = featuredPackages.find(p => p.name === name);
    const price = product ? salePrice(product) : pack ? Number(String(pack.price).replace(/[^0-9]/g, '')) : 0;
    return { name, qty, price, cost: catalogueCost(product || pack), category: product ? product.category : pack ? 'Packages' : 'Other' };
  });
}

function orderMoney(order) {
  const lines = orderLines(order);
  const lineRevenue = lines.reduce((sum, line) => sum + line.qty * line.price, 0);
  const revenue = Number(order.subtotal) || lineRevenue || Math.max(0, Number(order.total || 0) - Number(order.dispatchFee || 0));
  const lineCost = lines.reduce((sum, line) => sum + line.qty * line.cost, 0);
  const cost = order.cost != null && Number(order.cost) > 0 ? Number(order.cost) : lineCost;
  return { lines, revenue, cost, profit: revenue - cost, fee: Number(order.dispatchFee || 0), missingCost: lines.some(line => !line.cost) };
}

function periodBounds(key, now = new Date()) {
  const today = new Date(now); today.setHours(0, 0, 0, 0);
  let start = null;
  if (key === 'week') { start = new Date(today); start.setDate(today.getDate() - (today.getDay() + 6) % 7); }
  if (key === 'month') start = new Date(today.getFullYear(), today.getMonth(), 1);
  if (key === '30d') { start = new Date(today); start.setDate(today.getDate() - 29); }
  if (key === 'year') start = new Date(today.getFullYear(), 0, 1);
  const end = new Date(now);
  if (!start) return { start, end, prevStart: null, prevEnd: null };
  const span = end - start;
  return { start, end, prevStart: new Date(start.getTime() - span - 1), prevEnd: new Date(start.getTime() - 1) };
}

function summarise(start, end) {
  const paid = orders.filter(isPaidOrder).filter(order => inRange(orderDate(order), start, end));
  const categories = {}, items = {}, customers = {}, days = {};
  let revenue = 0, cost = 0, fees = 0, missingCost = 0;
  paid.forEach(order => {
    const m = orderMoney(order);
    revenue += m.revenue; cost += m.cost; fees += m.fee; if (m.missingCost) missingCost++;
    m.lines.forEach(line => {
      const c = categories[line.category] || (categories[line.category] = { revenue: 0, profit: 0, units: 0 });
      c.revenue += line.qty * line.price; c.profit += line.qty * (line.price - line.cost); c.units += line.qty;
      const i = items[line.name] || (items[line.name] = { revenue: 0, profit: 0, units: 0, category: line.category });
      i.revenue += line.qty * line.price; i.profit += line.qty * (line.price - line.cost); i.units += line.qty;
    });
    const key = String(order.phone || '').replace(/\D/g, '') || order.name;
    const c = customers[key] || (customers[key] = { name: order.name, revenue: 0, orders: 0 });
    c.revenue += m.revenue; c.orders += 1;
    const d = order.deliveryDay || 'Wednesday';
    days[d] = (days[d] || 0) + m.revenue;
  });
  const spent = expenses.filter(e => inRange(expenseDate(e), start, end)).reduce((sum, e) => sum + Number(e.amount || 0), 0);
  const gross = revenue - cost;
  return { paid, revenue, cost, gross, fees, expenses: spent, net: gross - spent, margin: pct(gross, revenue), aov: paid.length ? Math.round(revenue / paid.length) : 0, missingCost, categories, items, customers, days };
}

function progressSeries(start, end) {
  let from = start;
  if (!from) {
    const dates = orders.filter(isPaidOrder).map(orderDate).concat(expenses.map(expenseDate)).filter(d => !isNaN(d));
    from = dates.length ? new Date(Math.min(...dates)) : new Date(end);
  }
  const monthly = (end - from) > 62 * 864e5;
  const buckets = [];
  const cursor = monthly ? new Date(from.getFullYear(), from.getMonth(), 1) : new Date(from.getFullYear(), from.getMonth(), from.getDate());
  while (cursor <= end && buckets.length < 400) {
    const next = monthly ? new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1) : new Date(cursor.getFullYear(), cursor.getMonth(), cursor.getDate() + 1);
    buckets.push({ start: new Date(cursor), end: new Date(next - 1), label: cursor.toLocaleDateString('en-GB', monthly ? { month: 'short', year: '2-digit' } : { day: 'numeric', month: 'short' }), value: 0 });
    cursor.setTime(next.getTime());
  }
  const place = (date, amount) => { const b = buckets.find(x => date >= x.start && date <= x.end); if (b) b.value += amount; };
  orders.filter(isPaidOrder).forEach(order => place(orderDate(order), orderMoney(order).profit));
  expenses.forEach(e => place(expenseDate(e), -Number(e.amount || 0)));
  return { buckets, monthly };
}

/* ---------- Views ---------- */

function progressChart(series) {
  const points = series.buckets;
  if (!points.length) return '';
  const W = 640, H = 220, L = 10, R = 10, T = 22, B = 30;
  let running = 0;
  const cumulative = points.map(p => (running += p.value));
  const values = [...points.map(p => p.value), ...cumulative, 0];
  const max = Math.max(...values), min = Math.min(...values), span = (max - min) || 1;
  const y = v => T + (max - v) / span * (H - T - B);
  const step = (W - L - R) / points.length;
  const cx = i => L + i * step + step / 2;
  const barWidth = Math.max(2, Math.min(24, step * 0.62));
  const bars = points.map((p, i) => { const y0 = y(0), y1 = y(p.value); return `<rect x="${(cx(i) - barWidth / 2).toFixed(1)}" y="${Math.min(y0, y1).toFixed(1)}" width="${barWidth.toFixed(1)}" height="${Math.max(1, Math.abs(y1 - y0)).toFixed(1)}" rx="2" class="${p.value < 0 ? 'bi-bar-neg' : 'bi-bar'}"><title>${esc(p.label)}: ${money(Math.round(p.value))}</title></rect>`; }).join('');
  const line = cumulative.map((v, i) => `${cx(i).toFixed(1)},${y(v).toFixed(1)}`).join(' ');
  const area = `M${cx(0).toFixed(1)},${y(0).toFixed(1)} L${line.replace(/ /g, ' L')} L${cx(points.length - 1).toFixed(1)},${y(0).toFixed(1)} Z`;
  const last = points.length - 1;
  const labelIdx = [...new Set([0, Math.floor(last / 2), last])];
  const labels = labelIdx.map(i => `<text x="${cx(i).toFixed(1)}" y="${H - 9}" text-anchor="${i === 0 ? 'start' : i === last ? 'end' : 'middle'}">${esc(points[i].label)}</text>`).join('');
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Profit per ${series.monthly ? 'month' : 'day'} and running total"><line x1="${L}" x2="${W - R}" y1="${y(0).toFixed(1)}" y2="${y(0).toFixed(1)}" class="bi-zero"/>${bars}<path d="${area}" class="bi-area"/><polyline points="${line}" class="bi-line"/><circle cx="${cx(last).toFixed(1)}" cy="${y(cumulative[last]).toFixed(1)}" r="4.5" class="bi-dot"/><text x="${W - R}" y="14" text-anchor="end" class="bi-chart-total">${esc(shortMoney(cumulative[last]))} running total</text>${labels}</svg>`;
}

function delta(current, previous) {
  if (previous === null || previous === undefined) return '';
  if (!previous) return current ? '<small class="bi-delta up">New this period</small>' : '';
  const change = Math.round((current - previous) / Math.abs(previous) * 100);
  return `<small class="bi-delta ${change >= 0 ? 'up' : 'down'}">${change >= 0 ? '▲' : '▼'} ${Math.abs(change)}% vs previous</small>`;
}

function rankRows(entries, valueKey, total, formatter, sub) {
  if (!entries.length) return '<div class="empty-catalogue"><b>Nothing to rank yet.</b><span>This fills in as paid orders come in.</span></div>';
  const top = Math.max(...entries.map(([, v]) => Math.abs(v[valueKey]))) || 1;
  return entries.map(([name, v], index) => `<article class="bi-rank"><div class="bi-rank-head"><b>${index + 1}. ${esc(name)}</b><strong class="${v[valueKey] < 0 ? 'neg' : ''}">${formatter(v[valueKey])}</strong></div><div class="bi-track"><i style="width:${Math.max(3, Math.abs(v[valueKey]) / top * 100)}%"></i></div><small>${sub(v, total)}</small></article>`).join('');
}

function targetCard() {
  const target = Number(localStorage.getItem(profitTargetKey) || 0);
  const now = new Date();
  const b = periodBounds('month', now);
  const s = summarise(b.start, b.end);
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const projected = Math.round(s.net / Math.max(1, now.getDate()) * daysInMonth);
  const form = `<form id="profit-target-form" class="bi-target-form"><input name="target" type="number" min="0" step="1000" value="${target || ''}" placeholder="Monthly profit goal (₦)" aria-label="Monthly profit goal"/><button class="button secondary" type="submit">${target ? 'Update goal' : 'Set goal'}</button></form>`;
  if (!target) return `<section class="bi-card"><div class="bi-card-head"><b>Monthly profit goal</b><span>${now.toLocaleDateString('en-GB', { month: 'long' })}</span></div><p class="bi-note">Set a goal to track progress and see if this month is on pace.</p>${form}</section>`;
  const done = Math.max(0, Math.min(100, pct(s.net, target)));
  const onPace = projected >= target;
  return `<section class="bi-card"><div class="bi-card-head"><b>Monthly profit goal</b><span>${now.toLocaleDateString('en-GB', { month: 'long' })}</span></div><div class="bi-goal"><strong class="${s.net < 0 ? 'neg' : ''}">${money(s.net)}</strong><span>of ${money(target)}</span></div><div class="bi-track big"><i style="width:${done}%"></i></div><p class="bi-note ${onPace ? 'good' : 'warn'}">${done}% reached · on pace for ${money(projected)} by month end${onPace ? ' — ahead of goal.' : ` — ${money(target - projected)} short at this pace.`}</p>${form}</section>`;
}

function businessReports() {
  const b = periodBounds(biPeriod);
  const s = summarise(b.start, b.end);
  const prev = b.start ? summarise(b.prevStart, b.prevEnd) : null;
  const periodLabel = biPeriods.find(([k]) => k === biPeriod)?.[1] || 'This month';
  const chips = `<div class="bi-periods" role="tablist">${biPeriods.map(([k, label]) => `<button data-bi-period="${k}" class="${k === biPeriod ? 'active' : ''}" role="tab" aria-selected="${k === biPeriod}">${label}</button>`).join('')}</div>`;
  const hero = `<section class="bi-hero"><span>Net profit · ${periodLabel.toLowerCase()}</span><b class="${s.net < 0 ? 'neg-light' : ''}">${money(s.net)}</b>${delta(s.net, prev?.net)}<div class="bi-hero-split"><div><small>Sales</small><strong>${money(s.revenue)}</strong></div><div><small>Gross profit</small><strong>${money(s.gross)}</strong></div><div><small>Margin</small><strong>${s.margin}%</strong></div></div></section>`;
  const kpis = `<div class="bi-kpis">
    <article class="bi-kpi"><span>Paid orders</span><b>${s.paid.length}</b>${delta(s.paid.length, prev?.paid.length)}</article>
    <article class="bi-kpi"><span>Average order</span><b>${money(s.aov)}</b>${delta(s.aov, prev?.aov)}</article>
    <article class="bi-kpi"><span>Cost of goods</span><b>${money(s.cost)}</b><small>${pct(s.cost, s.revenue)}% of sales</small></article>
    <article class="bi-kpi"><span>Expenses</span><b>${money(s.expenses)}</b><small>${pct(s.expenses, s.revenue)}% of sales</small></article>
  </div>`;
  const costWarning = s.missingCost ? `<div class="report-callout attention"><b>${s.missingCost} paid order${s.missingCost === 1 ? ' has' : 's have'} items without a cost price</b><span>Profit is overstated until you add cost prices in Catalogue › Edit.</span></div>` : '';
  const series = progressSeries(b.start, b.end);
  const chart = `<section class="bi-card"><div class="bi-card-head"><b>Profit progress</b><span>per ${series.monthly ? 'month' : 'day'} · after expenses</span></div><div class="bi-chart">${progressChart(series)}</div><div class="bi-legend"><span><i class="bar"></i>Profit that ${series.monthly ? 'month' : 'day'}</span><span><i class="line"></i>Running total</span></div></section>`;
  const categories = Object.entries(s.categories).sort((a, b2) => b2[1].profit - a[1].profit);
  const bestCategory = categories[0];
  const categoryCard = `<section class="bi-card"><div class="bi-card-head"><b>Best categories</b><span>by profit</span></div>${bestCategory ? `<p class="bi-note good">${esc(bestCategory[0])} leads with ${money(Math.round(bestCategory[1].profit))} profit — ${pct(bestCategory[1].revenue, s.revenue)}% of sales.</p>` : ''}${rankRows(categories, 'profit', s.revenue, v => money(Math.round(v)), v => `${money(Math.round(v.revenue))} sales · ${pct(v.profit, v.revenue)}% margin · ${v.units} unit${v.units === 1 ? '' : 's'}`)}</section>`;
  const byProfit = Object.entries(s.items).sort((a, b2) => b2[1].profit - a[1].profit).slice(0, 5);
  const byUnits = Object.entries(s.items).sort((a, b2) => b2[1].units - a[1].units).slice(0, 5);
  const productsCard = `<section class="bi-card"><div class="bi-card-head"><b>Most profitable products</b><span>top 5</span></div>${rankRows(byProfit, 'profit', s.revenue, v => money(Math.round(v)), v => `${esc(v.category)} · ${v.units} sold · ${pct(v.profit, v.revenue)}% margin`)}</section><section class="bi-card"><div class="bi-card-head"><b>Best sellers</b><span>by units</span></div>${rankRows(byUnits, 'units', 0, v => `${v} sold`, v => `${money(Math.round(v.revenue))} sales · ${money(Math.round(v.profit))} profit`)}</section>`;
  const topCustomers = Object.values(s.customers).sort((a, b2) => b2.revenue - a.revenue).slice(0, 3);
  const allCustomerOrders = {};
  orders.filter(isPaidOrder).forEach(o => { const k = String(o.phone || '').replace(/\D/g, '') || o.name; allCustomerOrders[k] = (allCustomerOrders[k] || 0) + 1; });
  const buyers = Object.keys(allCustomerOrders).length;
  const repeat = Object.values(allCustomerOrders).filter(n => n > 1).length;
  const customersCard = `<section class="bi-card"><div class="bi-card-head"><b>Customers</b><span>${Object.keys(s.customers).length} bought this period</span></div><div class="bi-mini-grid"><div><small>Repeat rate</small><strong>${pct(repeat, buyers)}%</strong></div><div><small>Repeat buyers</small><strong>${repeat} of ${buyers}</strong></div></div>${topCustomers.length ? topCustomers.map((c, i) => `<div class="bi-line-item"><span>${i + 1}. ${esc(c.name)}</span><b>${money(Math.round(c.revenue))}</b><small>${c.orders} order${c.orders === 1 ? '' : 's'}</small></div>`).join('') : '<p class="bi-note">Top customers appear once orders are paid.</p>'}</section>`;
  const active = products.filter(p => p.available !== false);
  const retail = active.reduce((sum, p) => sum + Number(p.stock || 0) * salePrice(p), 0);
  const costValue = active.reduce((sum, p) => sum + Number(p.stock || 0) * Number(p.costPrice || 0), 0);
  const noCost = active.filter(p => !Number(p.costPrice));
  const lowMargin = active.filter(p => Number(p.costPrice)).map(p => ({ p, m: pct(salePrice(p) - Number(p.costPrice), salePrice(p)) })).sort((a, b2) => a.m - b2.m).slice(0, 3);
  const stockCard = `<section class="bi-card"><div class="bi-card-head"><b>Stock & pricing health</b><span>${active.length} active product${active.length === 1 ? '' : 's'}</span></div><div class="bi-mini-grid three"><div><small>Stock at cost</small><strong>${shortMoney(costValue)}</strong></div><div><small>Stock at price</small><strong>${shortMoney(retail)}</strong></div><div><small>Profit in stock</small><strong>${shortMoney(retail - costValue)}</strong></div></div>${noCost.length ? `<p class="bi-note warn">${noCost.length} product${noCost.length === 1 ? ' has' : 's have'} no cost price: ${noCost.slice(0, 4).map(p => esc(p.name)).join(', ')}${noCost.length > 4 ? '…' : ''}</p>` : '<p class="bi-note good">Every active product has a cost price.</p>'}${lowMargin.map(({ p, m }) => `<div class="bi-line-item"><span>${esc(p.name)}</span><b class="${m < 15 ? 'neg' : ''}">${m}% margin</b><small>${money(salePrice(p))} price · ${money(Number(p.costPrice))} cost</small></div>`).join('')}</section>`;
  const awaiting = orders.filter(o => o.status === 'Awaiting payment');
  const toCollect = awaiting.reduce((sum, o) => sum + Number(o.balance ?? o.total ?? 0), 0);
  const plans = awaiting.filter(o => Number(o.installments || 0) > 1);
  const cashCard = `<section class="bi-card"><div class="bi-card-head"><b>Money to collect</b><span>right now</span></div><div class="bi-mini-grid three"><div><small>Unpaid orders</small><strong>${awaiting.length}</strong></div><div><small>To collect</small><strong>${shortMoney(toCollect)}</strong></div><div><small>Pay small small</small><strong>${plans.length} plan${plans.length === 1 ? '' : 's'}</strong></div></div><div class="bi-line-item"><span>Dispatch fees collected (${periodLabel.toLowerCase()})</span><b>${money(s.fees)}</b><small>Passed to dispatch — not counted as profit</small></div>${['Monday', 'Wednesday', 'Friday'].map(d => `<div class="bi-line-item"><span>${d} deliveries</span><b>${money(Math.round(s.days[d] || 0))}</b><small>${pct(s.days[d] || 0, s.revenue)}% of sales</small></div>`).join('')}</section>`;
  const periodExpenses = expenses.filter(e => inRange(expenseDate(e), b.start, b.end)).sort((a, b2) => b2.date.localeCompare(a.date));
  const byType = {};
  periodExpenses.forEach(e => { byType[e.category] = (byType[e.category] || 0) + Number(e.amount || 0); });
  const expenseCard = `<section class="bi-card" id="bi-expenses"><div class="bi-card-head"><b>Expenses</b><span>${money(s.expenses)} · ${periodLabel.toLowerCase()}</span></div>${Object.keys(byType).length ? `<div class="bi-chips">${Object.entries(byType).sort((a, b2) => b2[1] - a[1]).map(([k, v]) => `<span>${esc(k)} <b>${shortMoney(v)}</b></span>`).join('')}</div>` : ''}<form id="expense-form" class="form-card bi-expense-form"><div class="form-row"><label>Amount (₦)<input required name="amount" type="number" min="1" placeholder="0"/></label><label>Date<input required name="date" type="date" value="${dayKey(Date.now())}"/></label></div><label>Type<select required name="category">${expenseCategories.map(c => `<option>${c}</option>`).join('')}</select></label><label>Note <span>(optional)</span><input name="description" maxlength="120" placeholder="e.g. Fuel for Friday deliveries"/></label><button class="button" type="submit">Add expense</button></form><p class="bi-note">Record running costs here. Don't add stock purchases — those are already counted through cost prices.</p><div class="bi-expense-list">${periodExpenses.length ? periodExpenses.slice(0, 12).map(e => `<article><div><b>${esc(e.category)}${e.description ? ` · ${esc(e.description)}` : ''}</b><span>${new Date(`${e.date}T12:00:00`).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })}</span></div><strong>${money(e.amount)}</strong><button class="text-action archive" data-delete-expense="${esc(e.id)}" aria-label="Delete expense">Delete</button></article>`).join('') : '<div class="empty-catalogue"><b>No expenses recorded for this period.</b><span>Add transport, packaging, rent and other running costs.</span></div>'}</div></section>`;
  return `<section class="screen owner-screen reports-screen bi-screen">${ownerHeader('Business reports', 'Profit, progress and what is selling.')}${chips}${hero}${costWarning}${kpis}${targetCard()}${chart}${categoryCard}${productsCard}${customersCard}${stockCard}${cashCard}${expenseCard}</section>${ownerNav('reports')}`;
}
reportsV2 = businessReports;

/* ---------- Owner navigation, overview and More ---------- */

ownerNav = function (active) {
  const current = { 'owner-customers': 'owner-more', staff: 'owner-more', 'business-settings': 'owner-more' }[active] || active;
  return `<nav class="owner-nav" aria-label="Owner navigation">${[['owner-overview', 'Overview', 'home'], ['owner-orders', 'Orders', 'receipt'], ['owner-catalogue', 'Catalogue', 'box'], ['reports', 'Reports', 'chart'], ['owner-more', 'More', 'more']].map(([id, label, image]) => `<button class="${current === id ? 'active' : ''}" data-page="${id}">${icon(image)}<span>${label}</span></button>`).join('')}</nav>`;
};

const ownerOverviewBase = ownerOverview;
ownerOverview = function () {
  const b = periodBounds('month');
  const s = summarise(b.start, b.end);
  const today = summarise(new Date(new Date().setHours(0, 0, 0, 0)), new Date());
  const card = `<button class="bi-overview" data-page="reports"><div><span>Profit this month</span><b class="${s.net < 0 ? 'neg' : ''}">${money(s.net)}</b><small>${money(s.revenue)} sales · ${s.margin}% margin · today ${money(today.gross)} gross</small></div><i>›</i></button>`;
  return ownerOverviewBase().replace('<div class="section-head"><h2>Needs attention</h2>', `${card}<div class="section-head"><h2>Needs attention</h2>`);
};

const ownerMoreBusinessBase = ownerMoreV2;
ownerMoreV2 = function () {
  const extra = `<button data-page="owner-customers"><span>${icon('users')}</span><div><b>Customers</b><small>Everyone who has ordered</small></div><i>›</i></button><button data-bi-expenses><span>${icon('receipt')}</span><div><b>Expenses</b><small>Transport, packaging, rent and more</small></div><i>›</i></button>`;
  const setup = `<details class="sheet-setup"><summary>How to connect</summary><ol><li>In your Google Sheet open <b>Treasures › Show owner key</b>.</li><li>Paste the key above and tap <b>Save owner key</b>.</li><li>Tap <b>Load latest from Sheet</b>.</li></ol><p>The web-app link is set once in <b>config.js</b> by whoever manages the website.</p></details>`;
  return ownerMoreBusinessBase().replace('<div class="settings-list">', `<div class="settings-list">${extra}`).replace(/<details class="sheet-setup">[\s\S]*?<\/details>/, setup);
};

/* ---------- Events ---------- */

document.addEventListener('click', event => {
  const period = event.target.closest('[data-bi-period]');
  if (period) { event.preventDefault(); event.stopImmediatePropagation(); biPeriod = period.dataset.biPeriod; localStorage.setItem(biPeriodKey, biPeriod); render('reports'); return; }
  if (event.target.closest('[data-bi-expenses]')) { event.preventDefault(); event.stopImmediatePropagation(); render('reports'); requestAnimationFrame(() => document.getElementById('bi-expenses')?.scrollIntoView({ behavior: 'smooth', block: 'start' })); return; }
  const remove = event.target.closest('[data-delete-expense]');
  if (remove) {
    event.preventDefault(); event.stopImmediatePropagation();
    const expense = expenses.find(e => e.id === remove.dataset.deleteExpense);
    if (!expense || !window.confirm(`Delete this ${expense.category.toLowerCase()} expense of ${money(expense.amount)}?`)) return;
    expenses = expenses.filter(e => e.id !== expense.id); saveExpenses();
    syncToSheet('Expenses', { ...expenseRecord(expense), 'Deleted': true });
    showToast('Expense deleted.'); render('reports');
  }
}, true);

function expenseRecord(e) { return { 'Expense ID': e.id, 'Date': e.date, 'Category': e.category, 'Description': e.description || '', 'Amount (₦)': Number(e.amount || 0), 'Deleted': false, 'Updated at': new Date().toISOString() }; }

document.addEventListener('submit', event => {
  if (event.target.id === 'expense-form') {
    event.preventDefault(); event.stopImmediatePropagation();
    const f = event.target.elements;
    const expense = { id: `EXP-${Date.now()}`, date: f.date.value || dayKey(Date.now()), category: f.category.value, description: clean(f.description.value), amount: Math.max(0, Number(f.amount.value || 0)) };
    if (!expense.amount) { showToast('Enter the amount spent.'); return; }
    expenses.unshift(expense); saveExpenses(); syncToSheet('Expenses', expenseRecord(expense));
    showToast(`${expense.category} expense of ${money(expense.amount)} added.`);
    render('reports'); requestAnimationFrame(() => document.getElementById('bi-expenses')?.scrollIntoView({ block: 'start' }));
    return;
  }
  if (event.target.id === 'profit-target-form') {
    event.preventDefault(); event.stopImmediatePropagation();
    const target = Math.max(0, Number(event.target.elements.target.value || 0));
    if (target) localStorage.setItem(profitTargetKey, String(target)); else localStorage.removeItem(profitTargetKey);
    showToast(target ? `Monthly goal set to ${money(target)}.` : 'Monthly goal removed.');
    render('reports');
  }
}, true);

/* ---------- Keep the owner device in step with the Sheet ---------- */

const renderWithOwnerSync = render;
render = function (page) {
  const output = renderWithOwnerSync(page);
  if (ownerUnlocked && !ownerSyncedThisSession && sheetsEndpoint && sheetOwnerKey && location.hash === '#owner') {
    ownerSyncedThisSession = true;
    setTimeout(() => loadFromSheetSecure(), 0);
  }
  return output;
};
if (location.hash === '#owner' && ownerUnlocked) render('owner-overview');
