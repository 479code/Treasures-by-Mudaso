/*
 * Treasures by Mudaso — Google Sheets bridge
 * Paste into Extensions > Apps Script inside the Operations spreadsheet, then
 * Deploy > Manage deployments > Edit > Version: New version (execute as you, access: anyone).
 *
 * Script properties:
 *   OWNER_API_KEY         required — private owner phrase
 *   DISPATCH_FEE          optional — standard dispatch fee in naira
 *   PAYMENT_INSTRUCTIONS  optional — bank details shown to customers
 *   WHATSAPP              optional — support number, e.g. 2348012345678
 * The last three can also be saved from Owner > More > Dispatch & payment.
 */
const SPREADSHEET_ID = '1cdMp1MqIxg3tqdVWug12zCJjHQyYE0Q7-8beHznDLAI';
const PROPS = PropertiesService.getScriptProperties();
const SHEET_HEADERS = {
  Products: ['Product ID', 'Name', 'Category', 'Price (₦)', 'Sale price (₦)', 'On sale', 'Stock', 'Available', 'Image URL', 'Updated at'],
  Packages: ['Package ID', 'Name', 'Price (₦)', 'Label', 'Image URL', 'Contents', 'Components', 'Pay small small payments', 'Available', 'Updated at'],
  Orders: ['Order ID', 'Customer name', 'Phone', 'Address', 'Delivery day', 'Urgent', 'Items', 'Total (₦)', 'Payment status', 'Order status', 'Pay small small payments', 'Payments confirmed', 'Balance (₦)', 'Dispatch fee (₦)', 'Internal note', 'Substitution note', 'Packing checked', 'Created at', 'Updated at', 'Picked items'],
  Customers: ['Customer ID', 'Name', 'Phone', 'Address', 'Orders', 'Last order at'],
  Staff: ['Staff ID', 'Name', 'Role', 'Phone', 'Active', 'Updated at'],
  Images: ['Image ID', 'Image URL', 'File ID', 'Name', 'Uploaded at', 'Upload status']
};

function doGet(event) {
  const parameters = event && event.parameter || {};
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  const respond = data => parameters.callback ? jsonp(parameters.callback, data) : json(data);
  if (parameters.action === 'catalog') {
    const live = item => String(item.Available).toLowerCase() !== 'false';
    return respond({ products: readRows(book.getSheetByName('Products')).filter(live), packages: readRows(book.getSheetByName('Packages')).filter(live), settings: publicSettings() });
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
    // Only what a customer needs to follow their delivery — never address or internal notes.
    return respond(match ? { found: true, order: { id: match['Order ID'], items: match.Items || '', status: match['Order status'] || 'Awaiting payment', deliveryDay: match['Delivery day'] || '', urgent: String(match.Urgent).toLowerCase() === 'true' } } : { found: false });
  }
  return respond({ ok: false, error: 'Unknown action.' });
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents || '{}');
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (payload.action === 'createOrder') return withLock(() => createOrder(book, payload));
  if (payload.action === 'markPaid') return withLock(() => markPaid(book, payload));
  if (!ownerRequest(payload)) return json({ ok: false, error: 'Owner access required.' });
  if (payload.action === 'dump') {
    return json({ ok: true, products: readRows(book.getSheetByName('Products')), packages: readRows(book.getSheetByName('Packages')), orders: readRows(book.getSheetByName('Orders')), customers: readRows(book.getSheetByName('Customers')), staff: readRows(book.getSheetByName('Staff')) });
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
  return withLock(() => json(writeRecord(sheet, payload.sheet === 'Orders' ? guardPaymentFlag(sheet, payload.record) : payload.record)));
}

// Public: customers can create a new order only. Prices and stock come from the Sheet, not the browser.
function createOrder(book, payload) {
  const r = payload.record || {};
  const lines = Array.isArray(payload.lines) ? payload.lines : [];
  const id = clean(r['Order ID']);
  if (!/^TBM-[A-Z0-9]{6,20}$/.test(id) || !clean(r['Customer name']) || !clean(r.Phone) || !lines.length) return json({ ok: false, error: 'Please complete your name, phone number and bag.' });
  const orders = ensureSheet(book, 'Orders');
  if (findRow(orders, id)) return json({ ok: false, error: 'Order already exists.' });
  const priced = priceLines(book, lines);
  if (priced.error) return json({ ok: false, error: priced.error });
  const fee = publicSettings().deliveryFee;
  const total = priced.subtotal + fee;
  const now = new Date().toISOString();
  writeRecord(orders, {
    'Order ID': id, 'Customer name': clean(r['Customer name']), 'Phone': clean(r.Phone), 'Address': clean(r.Address), 'Items': clean(r.Items),
    'Delivery day': ['Monday', 'Wednesday', 'Friday'].includes(r['Delivery day']) ? r['Delivery day'] : 'Wednesday',
    'Urgent': String(r.Urgent) === 'true', 'Total (₦)': total, 'Balance (₦)': total, 'Dispatch fee (₦)': fee,
    'Pay small small payments': Math.min(4, Math.max(0, Number(r['Pay small small payments']) || 0)), 'Payments confirmed': 0,
    'Payment status': 'Awaiting payment', 'Order status': 'Awaiting payment', 'Internal note': '', 'Substitution note': '',
    'Packing checked': false, 'Picked items': '[]', 'Created at': now, 'Updated at': now
  });
  decrementStock(ensureSheet(book, 'Products'), priced.need);
  return json({ ok: true, orderId: id, total, dispatchFee: fee });
}

function priceLines(book, lines) {
  const products = readRows(book.getSheetByName('Products'));
  const packages = readRows(book.getSheetByName('Packages'));
  const live = row => String(row.Available).toLowerCase() !== 'false';
  const need = {};
  let subtotal = 0;
  for (const line of lines) {
    const qty = Math.max(1, Math.floor(Number(line.quantity) || 0));
    if (line.kind === 'product') {
      const p = products.find(row => String(row['Product ID']) === String(line.id));
      if (!p || !live(p)) return { error: 'An item in your bag is no longer available.' };
      const sale = String(p['On sale']).toLowerCase() === 'true' && Number(p['Sale price (₦)']) > 0;
      subtotal += qty * Number(sale ? p['Sale price (₦)'] : p['Price (₦)']);
      need[p['Product ID']] = (need[p['Product ID']] || 0) + qty;
    } else {
      const k = packages.find(row => String(row['Package ID']) === String(line.id));
      if (!k || !live(k)) return { error: 'A package in your bag is no longer available.' };
      subtotal += qty * Number(k['Price (₦)']);
      if (!Number(line.installments)) {
        let components = [];
        try { components = JSON.parse(k.Components || '[]'); } catch (e) {}
        components.forEach(c => { need[c.productId] = (need[c.productId] || 0) + qty * Number(c.quantity || 0); });
      }
    }
  }
  for (const [pid, q] of Object.entries(need)) {
    const p = products.find(row => String(row['Product ID']) === pid);
    if (!p || Number(p.Stock) < q) return { error: `Not enough stock: ${p ? p.Name : pid}` };
  }
  return { subtotal, need };
}

function decrementStock(sheet, need) {
  const col = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('Stock') + 1;
  if (!col) return;
  Object.entries(need).forEach(([pid, q]) => {
    const row = findRow(sheet, pid);
    if (row) { const cell = sheet.getRange(row, col); cell.setValue(Math.max(0, Number(cell.getValue()) - q)); }
  });
}

// Public: the customer who placed the order (matched by phone) can flag payment as sent.
function markPaid(book, payload) {
  const sheet = ensureSheet(book, 'Orders');
  const row = findRow(sheet, clean(payload.order));
  const phone = normalisePhone(payload.phone);
  if (!row || !phone) return json({ ok: false, error: 'Order not found.' });
  const order = readRow(sheet, row);
  if (normalisePhone(order.Phone) !== phone) return json({ ok: false, error: 'Order not found.' });
  if (order['Order status'] !== 'Awaiting payment') return json({ ok: false, error: 'This order is already confirmed.' });
  writeRecord(sheet, { ...order, 'Payment status': 'Customer marked paid', 'Updated at': new Date().toISOString() });
  return json({ ok: true });
}

// Stops an owner save from a stale device wiping a customer's "marked paid" flag.
function guardPaymentFlag(sheet, record) {
  const row = findRow(sheet, String(record['Order ID'] || ''));
  if (!row) return record;
  const current = readRow(sheet, row);
  const stillAwaiting = record['Order status'] === 'Awaiting payment' && current['Order status'] === 'Awaiting payment';
  const noNewConfirmation = Number(record['Payments confirmed'] || 0) <= Number(current['Payments confirmed'] || 0);
  return stillAwaiting && noNewConfirmation && current['Payment status'] === 'Customer marked paid' ? { ...record, 'Payment status': 'Customer marked paid' } : record;
}

function publicSettings() {
  return { deliveryFee: Number(PROPS.getProperty('DISPATCH_FEE') || 0), paymentInstructions: PROPS.getProperty('PAYMENT_INSTRUCTIONS') || '', whatsApp: PROPS.getProperty('WHATSAPP') || '' };
}

function clean(value) {
  return String(value ?? '').replace(/[<>"`]/g, '').replace(/'/g, '’').replace(/&/g, 'and').trim().slice(0, 500);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function upsertRecord(sheet, record) {
  return withLock(() => json(writeRecord(sheet, record)));
}

function writeRecord(sheet, record) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = headers.map(header => record[header] ?? '');
  const row = findRow(sheet, String(record[headers[0]] || ''));
  if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
  else sheet.appendRow(values);
  return { ok: true, row: row || sheet.getLastRow() };
}

function readRow(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  return Object.fromEntries(headers.map((header, index) => [header, values[index]]));
}

function findRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return null;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const offset = ids.findIndex(value => String(value) === id);
  return offset === -1 ? null : offset + 2;
}

function ownerRequest(values) {
  const expected = PROPS.getProperty('OWNER_API_KEY');
  return Boolean(expected && values && values.key && values.key === expected);
}

function normalisePhone(value) {
  return String(value || '').replace(/\D/g, '').replace(/^234/, '0').replace(/^([1-9])/, '0$1');
}

function productImageFolder() {
  const existingId = PROPS.getProperty('PRODUCT_IMAGES_FOLDER_ID');
  if (existingId) return DriveApp.getFolderById(existingId);
  const folders = DriveApp.getFoldersByName('Treasures by Mudaso Product Images');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Treasures by Mudaso Product Images');
  PROPS.setProperty('PRODUCT_IMAGES_FOLDER_ID', folder.getId());
  return folder;
}

// Run once from the editor to approve Drive photo uploads.
function authorizeProductImages() {
  const folder = productImageFolder();
  const check = folder.createFile('Treasures photo upload permission check.txt', 'This temporary file confirms that product photos can be uploaded and shared.');
  check.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  check.setTrashed(true);
  return 'Drive upload and public image sharing are authorised.';
}

// Run once from the editor to add new columns (e.g. Picked items) and format Phone as plain text.
function setupSheets() {
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SHEET_HEADERS).forEach(name => ensureSheet(book, name));
  const orders = book.getSheetByName('Orders');
  const phoneCol = orders.getRange(1, 1, 1, orders.getLastColumn()).getValues()[0].indexOf('Phone') + 1;
  if (phoneCol) orders.getRange(1, phoneCol, orders.getMaxRows(), 1).setNumberFormat('@');
  return 'Sheets ready.';
}

function readRows(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  return rows.filter(row => row[0]).map(row => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
}

function ensureSheet(book, name) {
  if (!name || !SHEET_HEADERS[name]) return null;
  let sheet = book.getSheetByName(name);
  if (sheet) {
    const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const missing = SHEET_HEADERS[name].filter(header => !current.includes(header));
    if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    return sheet;
  }
  sheet = book.insertSheet(name);
  sheet.getRange(1, 1, 1, SHEET_HEADERS[name].length).setValues([SHEET_HEADERS[name]]);
  sheet.setFrozenRows(1);
  return sheet;
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp(callback, data) {
  const safe = String(callback).replace(/[^\w$.]/g, '');
  return ContentService.createTextOutput(`${safe}(${JSON.stringify(data)})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
