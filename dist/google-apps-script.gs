/*
 * Treasures by Mudaso — Google Sheets bridge
 * Paste this into Extensions > Apps Script inside the Operations spreadsheet,
 * then deploy as a Web app (execute as you, access: anyone).
 */
const SPREADSHEET_ID = '1cdMp1MqIxg3tqdVWug12zCJjHQyYE0Q7-8beHznDLAI';
const SHEET_HEADERS = {
  Products: ['Product ID', 'Name', 'Category', 'Price (₦)', 'Sale price (₦)', 'On sale', 'Stock', 'Available', 'Image URL', 'Updated at'],
  Packages: ['Package ID', 'Name', 'Price (₦)', 'Label', 'Image URL', 'Contents', 'Components', 'Pay small small payments', 'Available', 'Updated at'],
  Orders: ['Order ID', 'Customer name', 'Phone', 'Address', 'Delivery day', 'Urgent', 'Items', 'Total (₦)', 'Payment status', 'Order status', 'Pay small small payments', 'Payments confirmed', 'Balance (₦)', 'Dispatch fee (₦)', 'Internal note', 'Substitution note', 'Packing checked', 'Created at', 'Updated at'],
  Customers: ['Customer ID', 'Name', 'Phone', 'Address', 'Orders', 'Last order at'],
  Staff: ['Staff ID', 'Name', 'Role', 'Phone', 'Active', 'Updated at'],
  Images: ['Image ID', 'Image URL', 'File ID', 'Name', 'Uploaded at', 'Upload status']
};

function doGet(event) {
  const parameters = event && event.parameter || {};
  const book = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (parameters.action === 'catalog') {
    const catalog = { products: readRows(book.getSheetByName('Products')).filter(item => String(item.Available).toLowerCase() !== 'false'), packages: readRows(book.getSheetByName('Packages')).filter(item => String(item.Available).toLowerCase() !== 'false') };
    return parameters.callback ? jsonp(parameters.callback, catalog) : json(catalog);
  }
  if (parameters.action === 'image') {
    const image = readRows(book.getSheetByName('Images')).find(item => String(item['Image ID']) === String(parameters.id));
    const output = { url: image ? image['Image URL'] : '', error: image ? image['Upload status'] : '' };
    return parameters.callback ? jsonp(parameters.callback, output) : json(output);
  }
  if (parameters.action === 'track') {
    const requestedId = String(parameters.order || '').trim();
    const requestedPhone = normalisePhone(parameters.phone);
    const match = readRows(book.getSheetByName('Orders')).find(item =>
      String(item['Order ID'] || '') === requestedId &&
      requestedPhone && normalisePhone(item.Phone) === requestedPhone
    );
    // Deliberately return only the details a customer needs to track their own
    // delivery — never their address, payment instructions or internal notes.
    const output = match ? {
      found: true,
      order: {
        id: match['Order ID'],
        items: match.Items || '',
        status: match['Order status'] || 'Awaiting payment',
        deliveryDay: match['Delivery day'] || '',
        urgent: String(match.Urgent).toLowerCase() === 'true'
      }
    } : { found: false };
    return parameters.callback ? jsonp(parameters.callback, output) : json(output);
  }
  if (!ownerRequest(parameters)) return parameters.callback ? jsonp(parameters.callback, { ok: false, error: 'Owner access required.' }) : json({ ok: false, error: 'Owner access required.' });
  const data = {
    products: readRows(book.getSheetByName('Products')),
    packages: readRows(book.getSheetByName('Packages')),
    orders: readRows(book.getSheetByName('Orders')),
    customers: readRows(book.getSheetByName('Customers')),
    staff: readRows(book.getSheetByName('Staff'))
  };
  const callback = parameters.callback;
  return callback ? jsonp(callback, data) : json(data);
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents || '{}');
  if (payload.action === 'createOrder') {
    const record = payload.record || {};
    if (!record['Order ID'] || !record['Customer name'] || !record.Phone || !record.Items) return json({ ok: false, error: 'Incomplete order.' });
    // A customer order must begin as unpaid, but retain the delivery fee shown
    // at checkout so owner reports and dispatch totals stay accurate.
    const safeOrder = {...record, 'Order status':'Awaiting payment', 'Payment status':'Awaiting payment', 'Dispatch fee (₦)':Number(record['Dispatch fee (₦)'] || 0), 'Internal note':'', 'Substitution note':'', 'Packing checked':false, 'Updated at':new Date().toISOString()};
    const orderSheet = ensureSheet(SpreadsheetApp.openById(SPREADSHEET_ID), 'Orders');
    return upsertRecord(orderSheet, safeOrder);
  }
  if (payload.action === 'uploadImage') {
    if (!ownerRequest(payload) || !payload.imageId || !payload.base64 || !payload.mimeType) return json({ ok: false, error: 'Owner access required.' });
    const imageSheet = ensureSheet(SpreadsheetApp.openById(SPREADSHEET_ID), 'Images');
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType, payload.name || 'product-image');
      const file = productImageFolder().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const imageRecord = {'Image ID':payload.imageId, 'Image URL':`https://drive.google.com/uc?export=view&id=${file.getId()}`, 'File ID':file.getId(), 'Name':payload.name || file.getName(), 'Uploaded at':new Date().toISOString(), 'Upload status':'ready'};
      return upsertRecord(imageSheet, imageRecord);
    } catch (error) {
      upsertRecord(imageSheet, {'Image ID':payload.imageId, 'Image URL':'', 'File ID':'', 'Name':payload.name || 'product-image', 'Uploaded at':new Date().toISOString(), 'Upload status':String(error.message || error)});
      return json({ ok: false, error: String(error.message || error) });
    }
  }
  if (!ownerRequest(payload) || payload.action !== 'upsert') return json({ ok: false, error: 'Owner access required.' });
  const sheet = ensureSheet(SpreadsheetApp.openById(SPREADSHEET_ID), payload.sheet);
  if (!sheet || !payload.record) return json({ ok: false, error: 'Invalid sheet or record.' });

  return upsertRecord(sheet, payload.record);
}

function upsertRecord(sheet, record) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idColumn = 0;
    const id = String(record[headers[idColumn]] || '');
    const values = headers.map(header => record[header] ?? '');
    const row = findRow(sheet, id);
    if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    return json({ ok: true, row: row || sheet.getLastRow() });
  } finally {
    lock.releaseLock();
  }
}

function findRow(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return null;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const offset = ids.findIndex(value => String(value) === id);
  return offset === -1 ? null : offset + 2;
}

function ownerRequest(values) {
  const expected = PropertiesService.getScriptProperties().getProperty('OWNER_API_KEY');
  return Boolean(expected && values && values.key && values.key === expected);
}

function normalisePhone(value) {
  return String(value || '').replace(/\D/g, '');
}

function productImageFolder() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty('PRODUCT_IMAGES_FOLDER_ID');
  if (existingId) return DriveApp.getFolderById(existingId);
  const folders = DriveApp.getFoldersByName('Treasures by Mudaso Product Images');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Treasures by Mudaso Product Images');
  properties.setProperty('PRODUCT_IMAGES_FOLDER_ID', folder.getId());
  return folder;
}

// Run once from the Apps Script editor to approve and verify Google Drive photo uploads.
// It creates then immediately trashes a small permission-check file.
function authorizeProductImages() {
  const folder = productImageFolder();
  const check = folder.createFile('Treasures photo upload permission check.txt', 'This temporary file confirms that product photos can be uploaded and shared.');
  check.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  check.setTrashed(true);
  return 'Drive upload and public image sharing are authorised.';
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
  return ContentService.createTextOutput(`${callback}(${JSON.stringify(data)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
