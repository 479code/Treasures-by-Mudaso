/**
 * Treasures by Mudaso — Google Sheets backend (v70)
 *
 * SETUP (new Google Sheet)
 *  1. In the Sheet: Extensions > Apps Script. Delete the starter code, paste this whole file, save.
 *  2. Reload the Sheet. A "Treasures" menu appears. Run Treasures > Set up sheets and approve access.
 *     This creates every tab, the three official packages and your private owner key.
 *  3. In Apps Script: Deploy > New deployment > type Web app.
 *     Execute as: Me. Who has access: Anyone. Deploy, then copy the Web app URL (ends in /exec).
 *  4. Paste that URL and this Sheet's URL into config.js in the website repo and push.
 *  5. In the app, open #owner > More, paste the owner key (Treasures > Show owner key) and save.
 *
 * After any later change to this file: Deploy > Manage deployments > Edit > Version: New version.
 * The /exec URL stays the same.
 */
const PROPS = PropertiesService.getScriptProperties();
const SHEET_HEADERS = {
  Products: ['Product ID', 'Name', 'Category', 'Price (₦)', 'Sale price (₦)', 'On sale', 'Cost price (₦)', 'Stock', 'Available', 'Image URL', 'Updated at'],
  Packages: ['Package ID', 'Name', 'Price (₦)', 'Cost price (₦)', 'Label', 'Image URL', 'Contents', 'Components', 'Pay small small payments', 'Available', 'Updated at'],
  Orders: ['Order ID', 'Customer name', 'Phone', 'Address', 'Delivery day', 'Urgent', 'Items', 'Subtotal (₦)', 'Cost (₦)', 'Dispatch fee (₦)', 'Total (₦)', 'Payment status', 'Order status', 'Pay small small payments', 'Payments confirmed', 'Balance (₦)', 'Internal note', 'Substitution note', 'Packing checked', 'Picked items', 'Lines', 'Created at', 'Paid at', 'Updated at'],
  Customers: ['Customer ID', 'Name', 'Phone', 'Address', 'Orders', 'Total ordered (₦)', 'First order at', 'Last order at'],
  Expenses: ['Expense ID', 'Date', 'Category', 'Description', 'Amount (₦)', 'Deleted', 'Updated at'],
  Staff: ['Staff ID', 'Name', 'Role', 'Phone', 'Active', 'Updated at'],
  Images: ['Image ID', 'Image URL', 'File ID', 'Name', 'Uploaded at', 'Upload status']
};
const PRIVATE_COLUMNS = ['Cost price (₦)', 'Components'];
const OFFICIAL_PACKAGES = [
  { 'Package ID': 'PKG-PREMIUM', 'Name': 'Premium grocery package', 'Price (₦)': 285000, 'Label': 'Most loved', 'Image URL': 'assets/premium-package.jpeg', 'Contents': '3 × Laziza spices (assorted), 1 × Soy sauce (1 litre), 6 × Laila basmati rice (1kg), 1 × Oats (1kg), 1 × Tetley tea (400 bags), 6 × assorted biscuits, 12 × pasta, 12 × spaghetti, 4 × John West tuna, 2 × 3-in-1 coffee.' },
  { 'Package ID': 'PKG-HEALTHY', 'Name': 'Healthy grocery package', 'Price (₦)': 787500, 'Label': 'Healthy choice', 'Image URL': 'assets/healthy-package.jpeg', 'Contents': '12 × cauliflower rice, 6 × brown rice, 2 × olive oil, 1 × demerara sugar (3kg), 1 × matcha and green tea, 1 × couscous, 2 × quinoa, 6 × gluten-free pasta, 1 × no-sugar muesli, 2 × oats (60 pack), 1 × Nando’s sauce, 1 × almond milk (pack of 8), 1 × organic honey, 1 × aloe vera (pack of 8), 1 × chia seeds, 2 × olives, 1 × pink salt, 2 × apple cider vinegar.' },
  { 'Package ID': 'PKG-ALL-INCLUSIVE', 'Name': 'All inclusive package', 'Price (₦)': 565000, 'Label': 'Delivery included', 'Image URL': 'assets/all-inclusive-package.jpeg', 'Contents': '2 cartons spaghetti (48 pieces), 18 × pasta, 6 × mayonnaise, 6 × ketchup, 6 × Laziza spices, 2 × soy sauce (500ml), 1 × basmati rice (10kg), 2 × oats (1kg), 1 × Tetley tea (400 bags), 2 × salad dressing, 4 × pasta sauces, 12 × biscuits.' }
];

/* ---------- Sheet menu & one-time setup ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Treasures')
    .addItem('Set up sheets', 'setup')
    .addItem('Show owner key', 'showOwnerKey')
    .addItem('Create a new owner key', 'rotateOwnerKey')
    .addToUi();
}

function setup() {
  const book = spreadsheet();
  Object.keys(SHEET_HEADERS).forEach(name => styleHeader(ensureSheet(book, name)));
  const orders = book.getSheetByName('Orders');
  const phoneCol = headerRow(orders).indexOf('Phone') + 1;
  orders.getRange(1, phoneCol, orders.getMaxRows(), 1).setNumberFormat('@');
  book.getSheetByName('Customers').getRange(1, 1, book.getSheetByName('Customers').getMaxRows(), 3).setNumberFormat('@');
  const packages = book.getSheetByName('Packages');
  OFFICIAL_PACKAGES.forEach(pack => {
    if (!findRow(packages, pack['Package ID'])) writeRecord(packages, { ...pack, 'Cost price (₦)': '', 'Components': '[]', 'Pay small small payments': 0, 'Available': true, 'Updated at': new Date().toISOString() });
  });
  const starter = book.getSheetByName('Sheet1');
  if (starter && starter.getLastRow() === 0 && book.getSheets().length > 1) book.deleteSheet(starter);
  productImageFolder();
  return notify(`Treasures is set up.\n\nYour private owner key:\n${ownerKey()}\n\nNext: Deploy > New deployment > Web app (Execute as: Me, Access: Anyone), then copy the /exec URL into config.js.`);
}

function showOwnerKey() { return notify(`Your private owner key:\n${ownerKey()}\n\nPaste it in the app: Owner > More > Private owner key.`); }

function rotateOwnerKey() {
  PROPS.deleteProperty('OWNER_API_KEY');
  return notify(`New owner key:\n${ownerKey()}\n\nThe old key stops working now. Update it on every owner device.`);
}

/* ---------- Web app: reads ---------- */

function doGet(event) {
  const parameters = event && event.parameter || {};
  const book = spreadsheet();
  const respond = data => parameters.callback ? jsonp(parameters.callback, data) : json(data);
  if (parameters.action === 'catalog') {
    const live = item => String(item.Available).toLowerCase() !== 'false';
    const publicRow = item => Object.fromEntries(Object.entries(item).filter(([key]) => !PRIVATE_COLUMNS.includes(key)));
    const packages = readRows(book.getSheetByName('Packages')).filter(live).map(item => ({ ...publicRow(item), Components: item.Components || '[]' }));
    return respond({ products: readRows(book.getSheetByName('Products')).filter(live).map(publicRow), packages, settings: publicSettings() });
  }
  if (parameters.action === 'image') {
    const image = readRows(book.getSheetByName('Images')).find(item => String(item['Image ID']) === String(parameters.id));
    return respond({ url: image ? image['Image URL'] : '', error: image && image['Upload status'] !== 'ready' ? image['Upload status'] : '' });
  }
  if (parameters.action === 'track') {
    const requestedId = String(parameters.order || '').trim();
    const requestedPhone = normalisePhone(parameters.phone);
    const match = readRows(book.getSheetByName('Orders')).find(item =>
      String(item['Order ID'] || '') === requestedId && requestedPhone && normalisePhone(item.Phone) === requestedPhone);
    return respond(match ? { found: true, order: { id: match['Order ID'], items: match.Items || '', status: match['Order status'] || 'Awaiting payment', deliveryDay: match['Delivery day'] || '', urgent: String(match.Urgent).toLowerCase() === 'true' } } : { found: false });
  }
  return respond({ ok: true, service: 'Treasures by Mudaso', version: 70 });
}

/* ---------- Web app: writes ---------- */

function doPost(event) {
  const payload = JSON.parse(event.postData && event.postData.contents || '{}');
  const book = spreadsheet();
  if (payload.action === 'createOrder') return withLock(() => createOrder(book, payload));
  if (payload.action === 'markPaid') return withLock(() => markPaid(book, payload));
  if (!ownerRequest(payload)) return json({ ok: false, error: 'Owner access required.' });
  if (payload.action === 'dump') {
    const rows = name => readRows(book.getSheetByName(name));
    return json({ ok: true, products: rows('Products'), packages: rows('Packages'), orders: rows('Orders'), customers: rows('Customers'), staff: rows('Staff'), expenses: rows('Expenses') });
  }
  if (payload.action === 'saveSettings') {
    PROPS.setProperties({
      DISPATCH_FEE: String(Math.max(0, Number(payload.deliveryFee) || 0)),
      PAYMENT_INSTRUCTIONS: String(payload.paymentInstructions || '').replace(/[<>"`]/g, '').slice(0, 1000),
      WHATSAPP: String(payload.whatsApp || '').replace(/\D/g, '')
    });
    return json({ ok: true });
  }
  if (payload.action === 'uploadImage') {
    if (!payload.imageId || !payload.base64 || !payload.mimeType) return json({ ok: false, error: 'Invalid image.' });
    const imageSheet = ensureSheet(book, 'Images');
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType, payload.name || 'product-image');
      const file = productImageFolder().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return upsertRecord(imageSheet, { 'Image ID': payload.imageId, 'Image URL': `https://drive.google.com/uc?export=view&id=${file.getId()}`, 'File ID': file.getId(), 'Name': payload.name || file.getName(), 'Uploaded at': new Date().toISOString(), 'Upload status': 'ready' });
    } catch (error) {
      upsertRecord(imageSheet, { 'Image ID': payload.imageId, 'Name': payload.name || 'product-image', 'Uploaded at': new Date().toISOString(), 'Upload status': String(error.message || error) });
      return json({ ok: false, error: String(error.message || error) });
    }
  }
  if (payload.action !== 'upsert') return json({ ok: false, error: 'Unknown action.' });
  const sheet = ensureSheet(book, payload.sheet);
  if (!sheet || !payload.record) return json({ ok: false, error: 'Invalid sheet or record.' });
  return withLock(() => {
    let record = { ...payload.record };
    if (payload.sheet === 'Orders') record = guardPaymentFlag(sheet, record);
    if (payload.sheet === 'Products') record = applyStockDelta(sheet, record);
    delete record.__stockBase;
    return json(writeRecord(sheet, record));
  });
}

// Public. Prices, costs and stock come from the Sheet, never from the browser.
function createOrder(book, payload) {
  const r = payload.record || {};
  const lines = Array.isArray(payload.lines) ? payload.lines.slice(0, 60) : [];
  const id = clean(r['Order ID']);
  if (!/^TBM-[A-Z0-9]{6,20}$/.test(id) || !clean(r['Customer name']) || !normalisePhone(r.Phone) || !lines.length) return json({ ok: false, error: 'Please complete your name, phone number and bag.' });
  const orders = ensureSheet(book, 'Orders');
  if (findRow(orders, id)) return json({ ok: false, error: 'Order already exists.' });
  const priced = priceLines(book, lines);
  if (priced.error) return json({ ok: false, error: priced.error });
  const fee = publicSettings().deliveryFee;
  const total = priced.subtotal + fee;
  const now = new Date().toISOString();
  const order = {
    'Order ID': id, 'Customer name': clean(r['Customer name']), 'Phone': clean(r.Phone), 'Address': clean(r.Address),
    'Items': priced.lines.map(line => `${line.qty} × ${line.name}`).join(', '),
    'Delivery day': ['Monday', 'Wednesday', 'Friday'].includes(r['Delivery day']) ? r['Delivery day'] : 'Wednesday',
    'Urgent': String(r.Urgent) === 'true', 'Subtotal (₦)': priced.subtotal, 'Cost (₦)': priced.cost, 'Dispatch fee (₦)': fee,
    'Total (₦)': total, 'Balance (₦)': total, 'Payment status': 'Awaiting payment', 'Order status': 'Awaiting payment',
    'Pay small small payments': Math.min(4, Math.max(0, Number(r['Pay small small payments']) || 0)), 'Payments confirmed': 0,
    'Internal note': '', 'Substitution note': '', 'Packing checked': false, 'Picked items': '[]',
    'Lines': JSON.stringify(priced.lines), 'Created at': now, 'Paid at': '', 'Updated at': now
  };
  writeRecord(orders, order);
  decrementStock(ensureSheet(book, 'Products'), priced.need);
  recordCustomer(ensureSheet(book, 'Customers'), order);
  return json({ ok: true, orderId: id, total, dispatchFee: fee });
}

function priceLines(book, lines) {
  const products = readRows(book.getSheetByName('Products'));
  const packages = readRows(book.getSheetByName('Packages'));
  const live = row => String(row.Available).toLowerCase() !== 'false';
  const productCost = pid => Number((products.find(row => String(row['Product ID']) === String(pid)) || {})['Cost price (₦)'] || 0);
  const need = {};
  const priced = [];
  let subtotal = 0;
  let cost = 0;
  for (const line of lines) {
    const qty = Math.max(1, Math.min(999, Math.floor(Number(line.quantity) || 0)));
    if (line.kind === 'product') {
      const p = products.find(row => String(row['Product ID']) === String(line.id));
      if (!p || !live(p)) return { error: 'An item in your bag is no longer available.' };
      const sale = String(p['On sale']).toLowerCase() === 'true' && Number(p['Sale price (₦)']) > 0;
      const price = Number(sale ? p['Sale price (₦)'] : p['Price (₦)']);
      const unitCost = Number(p['Cost price (₦)'] || 0);
      subtotal += qty * price; cost += qty * unitCost;
      need[p['Product ID']] = (need[p['Product ID']] || 0) + qty;
      priced.push({ kind: 'product', id: p['Product ID'], name: p.Name, category: p.Category || 'Other', qty, price, cost: unitCost });
    } else {
      const k = packages.find(row => String(row['Package ID']) === String(line.id));
      if (!k || !live(k)) return { error: 'A package in your bag is no longer available.' };
      let components = [];
      try { components = JSON.parse(k.Components || '[]'); } catch (e) {}
      const price = Number(k['Price (₦)']);
      const unitCost = Number(k['Cost price (₦)'] || 0) || components.reduce((sum, c) => sum + Number(c.quantity || 0) * productCost(c.productId), 0);
      subtotal += qty * price; cost += qty * unitCost;
      if (!Number(line.installments)) components.forEach(c => { need[c.productId] = (need[c.productId] || 0) + qty * Number(c.quantity || 0); });
      priced.push({ kind: 'package', id: k['Package ID'], name: k.Name, category: 'Packages', qty, price, cost: unitCost, plan: Number(line.installments) || 0 });
    }
  }
  for (const [pid, q] of Object.entries(need)) {
    const p = products.find(row => String(row['Product ID']) === pid);
    if (!p || Number(p.Stock) < q) return { error: `Not enough stock: ${p ? p.Name : pid}` };
  }
  return { subtotal, cost, need, lines: priced };
}

function decrementStock(sheet, need) {
  const col = headerRow(sheet).indexOf('Stock') + 1;
  if (!col) return;
  Object.entries(need).forEach(([pid, q]) => {
    const row = findRow(sheet, pid);
    if (row) { const cell = sheet.getRange(row, col); cell.setValue(Math.max(0, Number(cell.getValue()) - q)); }
  });
}

function recordCustomer(sheet, order) {
  const phone = normalisePhone(order.Phone);
  const row = findRow(sheet, phone);
  const current = row ? readRow(sheet, row) : {};
  writeRecord(sheet, {
    'Customer ID': phone, 'Name': order['Customer name'], 'Phone': order.Phone, 'Address': order.Address,
    'Orders': Number(current.Orders || 0) + 1, 'Total ordered (₦)': Number(current['Total ordered (₦)'] || 0) + Number(order['Total (₦)'] || 0),
    'First order at': current['First order at'] || order['Created at'], 'Last order at': order['Created at']
  });
}

// Public. Only the customer who placed the order (matched by phone) can flag payment as sent.
function markPaid(book, payload) {
  const sheet = ensureSheet(book, 'Orders');
  const row = findRow(sheet, clean(payload.order));
  const phone = normalisePhone(payload.phone);
  if (!row || !phone) return json({ ok: false, error: 'Order not found.' });
  const order = readRow(sheet, row);
  if (normalisePhone(order.Phone) !== phone) return json({ ok: false, error: 'Order not found.' });
  if (order['Order status'] !== 'Awaiting payment') return json({ ok: false, error: 'This order is already confirmed.' });
  writeRecord(sheet, { 'Order ID': order['Order ID'], 'Payment status': 'Customer marked paid', 'Updated at': new Date().toISOString() });
  return json({ ok: true });
}

// Stops an owner save from a stale device wiping the customer's "marked paid" flag.
function guardPaymentFlag(sheet, record) {
  const row = findRow(sheet, String(record['Order ID'] || ''));
  if (!row) return record;
  const current = readRow(sheet, row);
  const stillAwaiting = record['Order status'] === 'Awaiting payment' && current['Order status'] === 'Awaiting payment';
  const noNewConfirmation = Number(record['Payments confirmed'] || 0) <= Number(current['Payments confirmed'] || 0);
  return stillAwaiting && noNewConfirmation && current['Payment status'] === 'Customer marked paid' ? { ...record, 'Payment status': 'Customer marked paid' } : record;
}

// Owner edits send the stock they last saw; apply only the change, so customer sales in between aren't lost.
function applyStockDelta(sheet, record) {
  const row = findRow(sheet, String(record['Product ID'] || ''));
  if (!row || record.__stockBase === undefined || record.__stockBase === null || record.__stockBase === '') return record;
  const current = Number(readRow(sheet, row).Stock || 0);
  return { ...record, Stock: Math.max(0, current + Number(record.Stock || 0) - Number(record.__stockBase || 0)) };
}

/* ---------- Helpers ---------- */

function spreadsheet() {
  const id = PROPS.getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function publicSettings() {
  return { deliveryFee: Number(PROPS.getProperty('DISPATCH_FEE') || 0), paymentInstructions: PROPS.getProperty('PAYMENT_INSTRUCTIONS') || '', whatsApp: PROPS.getProperty('WHATSAPP') || '' };
}

function ownerKey() {
  let key = PROPS.getProperty('OWNER_API_KEY');
  if (!key) { key = `tbm-${Utilities.getUuid().replace(/-/g, '')}`; PROPS.setProperty('OWNER_API_KEY', key); }
  return key;
}

function ownerRequest(values) {
  const expected = PROPS.getProperty('OWNER_API_KEY');
  return Boolean(expected && values && values.key && values.key === expected);
}

function clean(value) {
  return String(value ?? '').replace(/[<>"`]/g, '').replace(/'/g, '’').replace(/&/g, 'and').trim().slice(0, 500);
}

function normalisePhone(value) {
  return String(value || '').replace(/\D/g, '').replace(/^234/, '0').replace(/^([1-9])/, '0$1');
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function upsertRecord(sheet, record) { return withLock(() => json(writeRecord(sheet, record))); }

// Updates only the columns present in the record; everything else in the row is kept.
function writeRecord(sheet, record) {
  const headers = headerRow(sheet);
  const row = findRow(sheet, String(record[headers[0]] || ''));
  const current = row ? sheet.getRange(row, 1, 1, headers.length).getValues()[0] : [];
  const values = headers.map((header, index) => header in record ? (record[header] ?? '') : (current[index] ?? ''));
  if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
  else sheet.appendRow(values);
  return { ok: true, row: row || sheet.getLastRow() };
}

function headerRow(sheet) { return sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]; }

function readRow(sheet, row) {
  const headers = headerRow(sheet);
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  return Object.fromEntries(headers.map((header, index) => [header, values[index]]));
}

function findRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return null;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const offset = ids.findIndex(value => String(value) === String(id));
  return offset === -1 ? null : offset + 2;
}

function readRows(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.filter(row => row[0] !== '' && row[0] !== null).map(row => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
}

function ensureSheet(book, name) {
  if (!name || !SHEET_HEADERS[name]) return null;
  let sheet = book.getSheetByName(name);
  if (sheet) {
    const current = sheet.getLastColumn() ? headerRow(sheet) : [];
    const missing = SHEET_HEADERS[name].filter(header => !current.includes(header));
    if (missing.length) sheet.getRange(1, current.filter(String).length + 1, 1, missing.length).setValues([missing]);
    return sheet;
  }
  sheet = book.insertSheet(name);
  sheet.getRange(1, 1, 1, SHEET_HEADERS[name].length).setValues([SHEET_HEADERS[name]]);
  sheet.setFrozenRows(1);
  return sheet;
}

function styleHeader(sheet) {
  const width = sheet.getLastColumn();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, width).setFontWeight('bold').setFontColor('#ffffff').setBackground('#805d37');
  sheet.autoResizeColumns(1, width);
}

function productImageFolder() {
  const existingId = PROPS.getProperty('PRODUCT_IMAGES_FOLDER_ID');
  if (existingId) { try { return DriveApp.getFolderById(existingId); } catch (e) {} }
  const folders = DriveApp.getFoldersByName('Treasures by Mudaso Product Images');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Treasures by Mudaso Product Images');
  PROPS.setProperty('PRODUCT_IMAGES_FOLDER_ID', folder.getId());
  return folder;
}

function notify(message) {
  Logger.log(message);
  try { SpreadsheetApp.getUi().alert(message); } catch (e) {}
  return message;
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp(callback, data) {
  const safe = String(callback).replace(/[^\w$.]/g, '');
  return ContentService.createTextOutput(`${safe}(${JSON.stringify(data)})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
