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
 *
 * ORDER LIFECYCLE:
 * Quotation Requested -> Quotation Sent -> Order Confirmed -> Measurement
 * Taken -> Work In Progress -> Ready for Collection -> Delivered (or
 * Cancelled). A customer rejecting a quotation (respondToQuotation) sends
 * the order back to "Quotation Requested" for a revised quote. Each status
 * change also writes a Notification row tagged with that same status string
 * as its Type and the OrderID, so a per-order timeline can be rebuilt by
 * filtering Notifications on OrderID instead of keeping a separate log.
 *
 * NOTE ON RE-RUNNING setupSheets(): the Quotations tab's columns changed
 * from a single Description/Amount pair to itemized StitchingCost /
 * FabricCost / AdditionalCost / TotalAmount / Notes. Re-running setupSheets()
 * only rewrites the header row, so any pre-existing Quotations rows from an
 * earlier version of this script will end up misaligned under the new
 * headers — clear or migrate that tab's data by hand before redeploying if
 * it has real rows in it.
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
  Orders: ['OrderID', 'UserID', 'Garment', 'Purpose', 'Date', 'TimeSlot', 'Notes', 'Status', 'CreatedAt', 'UpdatedAt', 'PhotoURLs'],
  Quotations: ['QuotationID', 'OrderID', 'UserID', 'StitchingCost', 'FabricCost', 'AdditionalCost', 'TotalAmount', 'Notes', 'Status', 'CreatedAt', 'UpdatedAt'],
  Measurements: ['MeasurementID', 'UserID', 'Garment', 'MeasurementsJSON', 'CreatedAt', 'UpdatedAt'],
  Notifications: ['NotificationID', 'UserID', 'Message', 'Type', 'IsRead', 'CreatedAt', 'OrderID'],
  Admin: ['AdminID', 'Name', 'Phone', 'Email', 'PasswordHash', 'Salt', 'CreatedAt']
};

// Customer-facing order lifecycle. Order status doubles as each timeline
// notification's Type, so the dashboard can rebuild a per-order timeline by
// filtering Notifications on OrderID instead of storing a separate log.
var ORDER_STATUSES = [
  'Quotation Requested', 'Quotation Sent', 'Order Confirmed',
  'Measurement Taken', 'Work In Progress', 'Ready for Collection',
  'Delivered', 'Cancelled'
];

var ACTIONS = {
  createUser: createUser,
  loginUser: loginUser,
  createOrder: createOrder,
  getOrdersByUser: getOrdersByUser,
  getAllOrders: getAllOrders,
  updateOrderStatus: updateOrderStatus,
  saveQuotation: saveQuotation,
  respondToQuotation: respondToQuotation,
  getQuotationsByUser: getQuotationsByUser,
  saveMeasurements: saveMeasurements,
  getMeasurementsByUser: getMeasurementsByUser,
  getNotifications: getNotifications,
  markNotificationRead: markNotificationRead
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

  var order = setOrderStatus_(orderId, status);
  if (!order) {
    return fail_('Order not found');
  }

  createNotification_(order.UserID, 'Your order status has been updated to "' + status + '".', status, orderId);
  return ok_({ orderId: orderId, status: status });
}

/** Writes a new Status + UpdatedAt onto an order row. Returns the pre-update row, or null if not found. */
function setOrderStatus_(orderId, status) {
  var found = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!found) return null;

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.ORDERS);
    var headers = SHEET_HEADERS.Orders;
    sheet.getRange(found.rowIndex, headers.indexOf('Status') + 1).setValue(status);
    sheet.getRange(found.rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
  });

  return found.row;
}

// ---------- Quotations ----------

function saveQuotation(data) {
  var orderId = (data.orderId || '').trim();
  var userId = (data.userId || '').trim();
  var stitchingCost = Number(data.stitchingCost) || 0;
  var fabricCost = Number(data.fabricCost) || 0;
  var additionalCost = Number(data.additionalCost) || 0;

  if (!orderId || !userId) {
    return fail_('orderId and userId are required');
  }
  if (!findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId)) {
    return fail_('Order not found');
  }

  var totalAmount = stitchingCost + fabricCost + additionalCost;
  var quotationId = genId_('QUO');
  var now = new Date();
  appendRow_(getSheet_(SHEET_NAMES.QUOTATIONS), SHEET_HEADERS.Quotations, {
    QuotationID: quotationId,
    OrderID: orderId,
    UserID: userId,
    StitchingCost: stitchingCost,
    FabricCost: fabricCost,
    AdditionalCost: additionalCost,
    TotalAmount: totalAmount,
    Notes: data.notes || '',
    Status: 'Pending',
    CreatedAt: now,
    UpdatedAt: now
  });

  setOrderStatus_(orderId, 'Quotation Sent');
  createNotification_(userId, 'A quotation of ₹' + totalAmount + ' has been sent for your order.', 'Quotation Sent', orderId);
  return ok_({ quotationId: quotationId, totalAmount: totalAmount });
}

function respondToQuotation(data) {
  var userId = (data.userId || '').trim();
  var orderId = (data.orderId || '').trim();
  var quotationId = (data.quotationId || '').trim();
  var decision = (data.decision || '').trim().toLowerCase();

  if (!userId || !orderId || !quotationId || (decision !== 'approve' && decision !== 'reject')) {
    return fail_('userId, orderId, quotationId and a decision of "approve" or "reject" are required');
  }

  var order = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!order || order.row.UserID !== userId) {
    return fail_('Order not found');
  }
  if (order.row.Status !== 'Quotation Sent') {
    return fail_('This order is not awaiting a quotation response');
  }

  var quotation = findById_(SHEET_NAMES.QUOTATIONS, 'QuotationID', quotationId);
  if (!quotation || quotation.row.OrderID !== orderId) {
    return fail_('Quotation not found');
  }

  var quotationStatus = decision === 'approve' ? 'Approved' : 'Rejected';
  var newOrderStatus = decision === 'approve' ? 'Order Confirmed' : 'Quotation Requested';

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.QUOTATIONS);
    var headers = SHEET_HEADERS.Quotations;
    sheet.getRange(quotation.rowIndex, headers.indexOf('Status') + 1).setValue(quotationStatus);
    sheet.getRange(quotation.rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
  });

  setOrderStatus_(orderId, newOrderStatus);
  createNotification_(
    userId,
    decision === 'approve'
      ? 'You approved the quotation — your order is now confirmed.'
      : 'You rejected the quotation. We\'ll follow up with a revised one.',
    newOrderStatus,
    orderId
  );

  return ok_({ orderId: orderId, quotationId: quotationId, status: newOrderStatus });
}

function getQuotationsByUser(data) {
  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }

  var quotations = sheetToObjects_(getSheet_(SHEET_NAMES.QUOTATIONS))
    .filter(function (row) { return row.UserID === userId; })
    .sort(byCreatedAtDesc_);

  return ok_({ quotations: quotations });
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

function getMeasurementsByUser(data) {
  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }

  var measurements = sheetToObjects_(getSheet_(SHEET_NAMES.MEASUREMENTS))
    .filter(function (row) { return row.UserID === userId; })
    .map(function (row) {
      var parsed = {};
      try { parsed = JSON.parse(row.MeasurementsJSON || '{}'); } catch (err) { parsed = {}; }
      return { garment: row.Garment, measurements: parsed, updatedAt: row.UpdatedAt };
    });

  return ok_({ measurements: measurements });
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

function markNotificationRead(data) {
  var userId = (data.userId || '').trim();
  var notificationId = (data.notificationId || '').trim();

  if (!userId || !notificationId) {
    return fail_('userId and notificationId are required');
  }

  var found = findById_(SHEET_NAMES.NOTIFICATIONS, 'NotificationID', notificationId);
  if (!found || found.row.UserID !== userId) {
    return fail_('Notification not found');
  }

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.NOTIFICATIONS);
    var headers = SHEET_HEADERS.Notifications;
    sheet.getRange(found.rowIndex, headers.indexOf('IsRead') + 1).setValue(true);
  });

  return ok_({ notificationId: notificationId });
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
    Status: 'Quotation Requested',
    CreatedAt: now,
    UpdatedAt: now,
    PhotoURLs: ''
  };
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.ORDERS), SHEET_HEADERS.Orders, record);
  });
  createNotification_(userId, 'Your order for ' + garment + ' has been received.', 'Quotation Requested', record.OrderID);
  return record;
}

function createNotification_(userId, message, type, orderId) {
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.NOTIFICATIONS), SHEET_HEADERS.Notifications, {
      NotificationID: genId_('NOT'),
      UserID: userId,
      Message: message,
      Type: type || 'general',
      IsRead: false,
      CreatedAt: new Date(),
      OrderID: orderId || ''
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
