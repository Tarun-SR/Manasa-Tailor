/**
 * Manasa Tailor — Backend API (v2)
 *
 * Google Apps Script + Google Sheets backend for the boutique site: user
 * accounts, orders, quotations, measurements and notifications, plus a
 * simple admin gate. No external services — everything lives in one Sheet.
 *
 * SETUP:
 * 1. Create a new Google Sheet (e.g. "Manasa Tailor Data").
 * 2. Extensions -> Apps Script. Delete any starter code and paste this
 *    whole file in.
 * 3. Select the "setupSheets" function in the toolbar dropdown and click
 *    Run once. This creates the Users / Orders / Quotations / Measurements
 *    / Notifications / Admin tabs with header rows. Safe to re-run — it
 *    skips tabs that already exist.
 * 4. Open the "Admin" tab and add at least one row by hand (AdminID can be
 *    anything you like, e.g. "admin1") so getAllOrders / updateOrderStatus
 *    have someone to authorize. There's no signup flow for admins on
 *    purpose — that's a follow-up issue.
 * 5. Deploy -> New deployment -> type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the Web App URL and wire it into index.html (see README).
 *
 * Every time you edit this script, create a NEW deployment (or "Manage
 * deployments" -> edit -> new version) for the change to go live — the URL
 * stays the same either way.
 *
 * CALLING THE API:
 * Both doGet and doPost accept an `action` field (query param for GET,
 * JSON body for POST) naming one of the functions below, plus whatever
 * params that function needs. Every response is JSON:
 *   { success: true,  data: {...} }
 *   { success: false, error: "message" }
 *
 * POST bodies with no `action` are treated as legacy order-enquiry
 * submissions from the original single-sheet version of this script (see
 * legacyOrderEnquiry_), so the live booking form keeps working unchanged.
 */

var SHEET_NAMES = {
  USERS: 'Users',
  ORDERS: 'Orders',
  QUOTATIONS: 'Quotations',
  MEASUREMENTS: 'Measurements',
  NOTIFICATIONS: 'Notifications',
  ADMIN: 'Admin'
};

var SHEET_HEADERS = {
  Users: ['UserID', 'Name', 'Phone', 'Email', 'PasswordHash', 'Salt', 'CreatedAt', 'City'],
  Orders: ['OrderID', 'UserID', 'Garment', 'Purpose', 'Date', 'TimeSlot', 'Notes', 'Status', 'CreatedAt', 'UpdatedAt'],
  Quotations: ['QuotationID', 'OrderID', 'UserID', 'Description', 'Amount', 'Status', 'CreatedAt', 'UpdatedAt'],
  Measurements: ['MeasurementID', 'UserID', 'Garment', 'MeasurementsJSON', 'CreatedAt', 'UpdatedAt'],
  Notifications: ['NotificationID', 'UserID', 'Message', 'Type', 'IsRead', 'CreatedAt'],
  Admin: ['AdminID', 'Name', 'Phone', 'Email', 'PasswordHash', 'Salt', 'CreatedAt']
};

var ORDER_STATUSES = ['Pending', 'Confirmed', 'In Progress', 'Ready', 'Completed', 'Cancelled'];

var ACTIONS = {
  createUser: createUser,
  loginUser: loginUser,
  createOrder: createOrder,
  getOrdersByUser: getOrdersByUser,
  getAllOrders: getAllOrders,
  updateOrderStatus: updateOrderStatus,
  saveQuotation: saveQuotation,
  saveMeasurements: saveMeasurements,
  getNotifications: getNotifications
};

/** One-time setup: creates every tab from SHEET_HEADERS if it doesn't already exist. */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    var headers = SHEET_HEADERS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  });

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && defaultSheet.getLastColumn() <= 1) {
    ss.deleteSheet(defaultSheet);
  }
}

function doGet(e) {
  return handleRequest_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    body = (e && e.parameter) ? e.parameter : {};
  }

  if (!body.action) {
    return legacyOrderEnquiry_(body);
  }

  return handleRequest_(body);
}

function handleRequest_(params) {
  var action = params.action;
  var fn = action && ACTIONS[action];

  if (!fn) {
    return fail_('Unknown or missing action: ' + action);
  }

  try {
    return fn(params);
  } catch (err) {
    return fail_(err && err.message ? err.message : String(err));
  }
}

// ---------- Users ----------

function createUser(data) {
  var name = (data.name || '').trim();
  var phone = (data.phone || '').trim();
  var email = (data.email || '').trim().toLowerCase();
  var password = data.password || '';
  var city = (data.city || '').trim();

  if (!name || !phone || !email || !password) {
    return fail_('Name, mobile, email and password are required');
  }
  if (password.length < 8) {
    return fail_('Password must be at least 8 characters');
  }
  if (findUserByPhone_(phone)) {
    return fail_('An account with this mobile number already exists');
  }
  if (findUserByEmail_(email)) {
    return fail_('An account with this email already exists');
  }

  var user = createUserRecord_(name, phone, email, password, city);
  return ok_({ userId: user.UserID, name: user.Name, phone: user.Phone, email: user.Email, city: user.City });
}

function loginUser(data) {
  var email = (data.email || '').trim().toLowerCase();
  var password = data.password || '';

  if (!email || !password) {
    return fail_('Email and password are required');
  }

  var user = findUserByEmail_(email);
  if (!user || hashPassword_(password, user.Salt) !== user.PasswordHash) {
    return fail_('Incorrect email or password');
  }

  return ok_({ userId: user.UserID, name: user.Name, phone: user.Phone, email: user.Email, city: user.City });
}

// ---------- Orders ----------

function createOrder(data) {
  var userId = (data.userId || '').trim();
  var garment = (data.garment || '').trim();

  if (!userId || !garment) {
    return fail_('userId and garment are required');
  }
  if (!findById_(SHEET_NAMES.USERS, 'UserID', userId)) {
    return fail_('User not found');
  }

  var order = createOrderRecord_(userId, garment, data.purpose, data.date, data.slot, data.notes);
  return ok_({ orderId: order.OrderID });
}

function getOrdersByUser(data) {
  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }

  var orders = sheetToObjects_(getSheet_(SHEET_NAMES.ORDERS))
    .filter(function (row) { return row.UserID === userId; })
    .sort(byCreatedAtDesc_);

  return ok_({ orders: orders });
}

function getAllOrders(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var orders = sheetToObjects_(getSheet_(SHEET_NAMES.ORDERS)).sort(byCreatedAtDesc_);
  return ok_({ orders: orders });
}

function updateOrderStatus(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var orderId = (data.orderId || '').trim();
  var status = (data.status || '').trim();

  if (!orderId || !status) {
    return fail_('orderId and status are required');
  }
  if (ORDER_STATUSES.indexOf(status) === -1) {
    return fail_('Invalid status. Must be one of: ' + ORDER_STATUSES.join(', '));
  }

  var found = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!found) {
    return fail_('Order not found');
  }

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.ORDERS);
    var headers = SHEET_HEADERS.Orders;
    sheet.getRange(found.rowIndex, headers.indexOf('Status') + 1).setValue(status);
    sheet.getRange(found.rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
  });

  createNotification_(found.row.UserID, 'Your order status has been updated to "' + status + '".', 'status');
  return ok_({ orderId: orderId, status: status });
}

// ---------- Quotations ----------

function saveQuotation(data) {
  var orderId = (data.orderId || '').trim();
  var userId = (data.userId || '').trim();
  var amount = data.amount;

  if (!orderId || !userId || amount === undefined || amount === null || amount === '') {
    return fail_('orderId, userId and amount are required');
  }
  if (!findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId)) {
    return fail_('Order not found');
  }

  var quotationId = genId_('QUO');
  var now = new Date();
  appendRow_(getSheet_(SHEET_NAMES.QUOTATIONS), SHEET_HEADERS.Quotations, {
    QuotationID: quotationId,
    OrderID: orderId,
    UserID: userId,
    Description: data.description || '',
    Amount: amount,
    Status: 'Pending',
    CreatedAt: now,
    UpdatedAt: now
  });

  createNotification_(userId, 'A quotation of ₹' + amount + ' has been added to your order.', 'quotation');
  return ok_({ quotationId: quotationId });
}

// ---------- Measurements ----------

function saveMeasurements(data) {
  var userId = (data.userId || '').trim();
  var garment = (data.garment || '').trim();
  var measurements = data.measurements;

  if (!userId || !garment || !measurements || typeof measurements !== 'object') {
    return fail_('userId, garment and a measurements object are required');
  }
  if (!findById_(SHEET_NAMES.USERS, 'UserID', userId)) {
    return fail_('User not found');
  }

  var sheet = getSheet_(SHEET_NAMES.MEASUREMENTS);
  var headers = SHEET_HEADERS.Measurements;
  var rows = sheetToObjects_(sheet);
  var existingIndex = rows.findIndex(function (row) {
    return row.UserID === userId && row.Garment === garment;
  });
  var now = new Date();
  var json = JSON.stringify(measurements);

  withLock_(function () {
    if (existingIndex === -1) {
      appendRow_(sheet, headers, {
        MeasurementID: genId_('MSR'),
        UserID: userId,
        Garment: garment,
        MeasurementsJSON: json,
        CreatedAt: now,
        UpdatedAt: now
      });
    } else {
      var rowIndex = existingIndex + 2; // header row + 0-based index
      sheet.getRange(rowIndex, headers.indexOf('MeasurementsJSON') + 1).setValue(json);
      sheet.getRange(rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(now);
    }
  });

  return ok_({ saved: true });
}

// ---------- Notifications ----------

function getNotifications(data) {
  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }

  var sheet = getSheet_(SHEET_NAMES.NOTIFICATIONS);
  var headers = SHEET_HEADERS.Notifications;
  var rows = sheetToObjects_(sheet);

  if (data.markAllRead) {
    withLock_(function () {
      rows.forEach(function (row, i) {
        if (row.UserID === userId && !row.IsRead) {
          sheet.getRange(i + 2, headers.indexOf('IsRead') + 1).setValue(true);
        }
      });
    });
    rows.forEach(function (row) {
      if (row.UserID === userId) row.IsRead = true;
    });
  }

  var notifications = rows
    .filter(function (row) {
      if (row.UserID !== userId) return false;
      if (data.unreadOnly) return !row.IsRead;
      return true;
    })
    .sort(byCreatedAtDesc_);

  return ok_({ notifications: notifications });
}

// ---------- Legacy compatibility (pre-v2 booking form) ----------

function legacyOrderEnquiry_(data) {
  var name = (data.name || '').trim();
  var phone = (data.phone || '').trim();
  var garment = (data.garment || '').trim();

  if (!phone || !garment) {
    return fail_('phone and garment are required');
  }

  var user = findUserByPhone_(phone);
  if (!user) {
    user = createUserRecord_(name || 'Guest', phone, '', Utilities.getUuid(), '');
  }

  var order = createOrderRecord_(user.UserID, garment, data.purpose, data.date, data.slot, data.notes);
  return ok_({ orderId: order.OrderID });
}

// ---------- Shared record creation ----------

function createUserRecord_(name, phone, email, password, city) {
  var salt = makeSalt_();
  var userId = genId_('USR');
  var record = {
    UserID: userId,
    Name: name,
    Phone: phone,
    Email: email || '',
    PasswordHash: hashPassword_(password, salt),
    Salt: salt,
    CreatedAt: new Date(),
    City: city || ''
  };
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.USERS), SHEET_HEADERS.Users, record);
  });
  return record;
}

function createOrderRecord_(userId, garment, purpose, date, slot, notes) {
  var now = new Date();
  var record = {
    OrderID: genId_('ORD'),
    UserID: userId,
    Garment: garment,
    Purpose: purpose || '',
    Date: date || '',
    TimeSlot: slot || '',
    Notes: notes || '',
    Status: 'Pending',
    CreatedAt: now,
    UpdatedAt: now
  };
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.ORDERS), SHEET_HEADERS.Orders, record);
  });
  createNotification_(userId, 'Your order for ' + garment + ' has been received.', 'order');
  return record;
}

function createNotification_(userId, message, type) {
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.NOTIFICATIONS), SHEET_HEADERS.Notifications, {
      NotificationID: genId_('NOT'),
      UserID: userId,
      Message: message,
      Type: type || 'general',
      IsRead: false,
      CreatedAt: new Date()
    });
  });
}

// ---------- Admin gate ----------

function requireAdmin_(adminId) {
  adminId = (adminId || '').trim();
  if (!adminId || !findById_(SHEET_NAMES.ADMIN, 'AdminID', adminId)) {
    return fail_('Unauthorized: valid adminId required');
  }
  return null;
}

// ---------- Sheet helpers ----------

function getSheet_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found. Run setupSheets() once from the Apps Script editor.');
  }
  return sheet;
}

function sheetToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (header, i) { obj[header] = row[i]; });
    return obj;
  });
}

function appendRow_(sheet, headers, obj) {
  var row = headers.map(function (header) {
    return obj.hasOwnProperty(header) ? obj[header] : '';
  });
  sheet.appendRow(row);
}

function findById_(sheetName, idColumn, idValue) {
  var sheet = getSheet_(sheetName);
  var rows = sheetToObjects_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][idColumn] === idValue) {
      return { rowIndex: i + 2, row: rows[i] };
    }
  }
  return null;
}

function findUserByPhone_(phone) {
  var rows = sheetToObjects_(getSheet_(SHEET_NAMES.USERS));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Phone === phone) return rows[i];
  }
  return null;
}

function findUserByEmail_(email) {
  var rows = sheetToObjects_(getSheet_(SHEET_NAMES.USERS));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Email && rows[i].Email.toLowerCase() === email) return rows[i];
  }
  return null;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}

function byCreatedAtDesc_(a, b) {
  return new Date(b.CreatedAt) - new Date(a.CreatedAt);
}

// ---------- ID + password helpers ----------

function genId_(prefix) {
  return prefix + '-' + Utilities.getUuid().split('-')[0].toUpperCase();
}

function makeSalt_() {
  return Utilities.getUuid();
}

function hashPassword_(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return digest.map(function (byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ---------- Response helpers ----------

function ok_(data) {
  return jsonOutput_({ success: true, data: data });
}

function fail_(message) {
  return jsonOutput_({ success: false, error: message });
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
