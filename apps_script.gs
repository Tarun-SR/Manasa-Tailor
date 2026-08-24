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
 * 3. Set SHEET_ID below to that Sheet's ID (the long ID in its URL, between
 *    /d/ and /edit). Every read/write goes through SpreadsheetApp.openById
 *    (SHEET_ID) rather than the "active" spreadsheet, so this script works
 *    the same whether it's bound to that Sheet or run as a standalone
 *    project — but it always operates on SHEET_ID's sheet specifically, not
 *    whichever one it happens to be pasted into.
 * 4. Select the "setupSheets" function in the toolbar dropdown and click
 *    Run once. This creates the Users / Orders / Quotations / Measurements
 *    / Notifications / Admin tabs with header rows. Safe to re-run — it
 *    skips tabs that already exist.
 * 5. Set up the one admin account: scroll to ADMIN_BOOTSTRAP_EMAIL /
 *    ADMIN_BOOTSTRAP_PASSWORD near the bottom of this file, fill them in,
 *    select "createAdminAccount" in the toolbar dropdown and click Run once.
 *    That's the email/password admin.html logs in with. Blank
 *    ADMIN_BOOTSTRAP_PASSWORD out again afterwards so it isn't sitting in
 *    the script in plain text — re-running with a new password just updates
 *    the same admin rather than creating a second one.
 * 6. Deploy -> New deployment -> type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Copy the Web App URL and wire it into index.html (see README).
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
 * GET IS THE PRIMARY TRANSPORT — every frontend page's shared callApi()
 * (in config.js) calls this over GET, not POST. Confirmed via live browser
 * testing (not just header inspection, which can't validate a browser's
 * actual CORS enforcement) that POST to this deployment's Web App URL
 * doesn't reliably make it back to the caller as readable JSON, while GET
 * does. Because of that, every value crosses as a URL string: object/array
 * params (measurements, photos) arrive JSON-stringified and are parsed back
 * out with parseMaybeJson_, and booleans arrive as the literal strings
 * "true"/"false" rather than real booleans, parsed with parseBool_ — a
 * naive `!!data.someBoolean` check would treat the string "false" as
 * truthy. Apply both helpers to any new action that takes an object, array,
 * or boolean param.
 *
 * The one exception is updateOrderStatus's photo upload (admin.html) — up
 * to 3 base64 images can exceed this endpoint's query-string ceiling
 * (confirmed failing somewhere between 8KB and 16KB), so that call stays on
 * a fire-and-forget no-cors POST instead, and the frontend can't read a
 * definitive success/failure from it — it just refetches state afterwards.
 *
 * ORDER LIFECYCLE:
 * Quotation Requested -> Quotation Sent -> Order Confirmed -> Measurement
 * Taken -> Work In Progress -> Ready for Collection -> Payment Done ->
 * Delivered (or Cancelled). A customer rejecting a quotation
 * (respondToQuotation) sends the order to "Negotiation" instead, along with
 * their reason (Orders.NegotiationNote) — saveQuotation can be called again
 * from there just like from Quotation Requested, sending a revised quote
 * and moving the order back to Quotation Sent. Ready for Collection ->
 * Payment Done requires a finalAmount and generates a PDF invoice (see
 * generateInvoice_) rather than being a plain status flip. Each status
 * change also writes a Notification row tagged with that same status string
 * as its Type and the OrderID, so a per-order timeline can be rebuilt by
 * filtering Notifications on OrderID instead of keeping a separate log.
 *
 * NOTE ON RE-RUNNING setupSheets(): the Quotations tab's columns changed
 * from a single Description/Amount pair to itemized StitchingCost /
 * FabricCost / AdditionalCost / TotalAmount / DeliveryTimeline / AdminNotes.
 * Re-running setupSheets() only rewrites the header row, so any pre-existing
 * Quotations rows from an earlier version of this script will end up
 * misaligned under the new headers — clear or migrate that tab's data by
 * hand before redeploying if it has real rows in it.
 *
 * ADMIN NOTES ARE INTERNAL: Quotations.AdminNotes is stripped out of
 * getQuotationsByUser's response before it's sent — never add it back there,
 * since that's the customer-facing endpoint. getQuotationsByOrder (admin-
 * only) is the one that includes it.
 *
 * ORDER PHOTOS: updateOrderStatus uploads photos to a "Manasa Tailor Order
 * Photos" Drive folder (auto-created) and stores their view URLs in
 * Orders.PhotoURLs — Drive access is part of Apps Script's default scopes
 * under "Execute as: Me", no extra API enablement needed.
 */

var SHEET_ID = '1TLMymlNJLjV3QSsMRSd3N7sTe7EyfBbPAn5jYiD-plw';

var SHEET_NAMES = {
  USERS: 'Users',
  ORDERS: 'Orders',
  QUOTATIONS: 'Quotations',
  MEASUREMENTS: 'Measurements',
  NOTIFICATIONS: 'Notifications',
  ADMIN: 'Admin',
  CLASSES: 'Classes',
  CLASS_ENROLLMENTS: 'ClassEnrollments'
};

var SHEET_HEADERS = {
  Users: ['UserID', 'Name', 'Phone', 'Email', 'PasswordHash', 'Salt', 'CreatedAt', 'City', 'Username'],
  Orders: ['OrderID', 'UserID', 'Garment', 'Purpose', 'Date', 'TimeSlot', 'Notes', 'Status', 'CreatedAt', 'UpdatedAt', 'PhotoURLs', 'HasMeasurements', 'NegotiationNote', 'AdvanceAmount', 'AdvancePaidAt', 'FinalAmountPaid', 'PaymentDoneAt', 'InvoiceURL'],
  Quotations: ['QuotationID', 'OrderID', 'UserID', 'StitchingCost', 'FabricCost', 'AdditionalCost', 'TotalAmount', 'DeliveryTimeline', 'AdminNotes', 'RequiredAdvance', 'Status', 'CreatedAt', 'UpdatedAt'],
  Measurements: ['MeasurementID', 'UserID', 'Garment', 'MeasurementsJSON', 'CreatedAt', 'UpdatedAt'],
  Notifications: ['NotificationID', 'UserID', 'Message', 'Type', 'IsRead', 'CreatedAt', 'OrderID'],
  Admin: ['AdminID', 'Name', 'Phone', 'Email', 'PasswordHash', 'Salt', 'CreatedAt'],
  Classes: ['ClassID', 'Title', 'Level', 'Description', 'Duration', 'TimingSlots', 'StartDate', 'TotalSeats', 'SeatsTaken', 'ContactNote', 'Fees', 'Status', 'CreatedAt', 'UpdatedAt'],
  ClassEnrollments: ['EnrollmentID', 'ClassID', 'UserID', 'CustomerName', 'Mobile', 'Email', 'PreferredLevel', 'PreferredTiming', 'Message', 'Status', 'AdminNotes', 'CreatedAt']
};

// Class lifecycle (admin-set on the Classes tab): Draft is being edited and
// never shown publicly; Published is what classes.html and index.html's
// teaser both fetch via getPublishedClasses; Closed stops taking new
// enrollments but the class record (and its enrollment history) stays.
var CLASS_STATUSES = ['Draft', 'Published', 'Closed'];

// Enrollment lifecycle (the "Enrollments" sub-tab in admin.html) — an
// enrollment is just an expression of interest, not an automatic seat
// booking, so this is a manual follow-up pipeline rather than something
// that writes back to Classes.SeatsTaken on its own.
var ENROLLMENT_STATUSES = ['New', 'Contacted', 'Confirmed', 'Enrolled', 'Completed', 'Not Interested'];

// Admin-driven status transitions (the "Status Update" dropdown in admin.html).
// Quotation Requested -> Quotation Sent happens only via saveQuotation, and
// Quotation Sent -> Order Confirmed / Negotiation only via respondToQuotation
// (the customer's approve/reject) — neither goes through this map.
// Ready for Collection -> Payment Done requires a finalAmount and generates
// the invoice (see updateOrderStatus) rather than being a plain status flip.
// 'Cancelled' is allowed as an admin override from any open status.
var VALID_STATUS_TRANSITIONS = {
  'Order Confirmed': ['Measurement Taken'],
  'Measurement Taken': ['Work In Progress'],
  'Work In Progress': ['Ready for Collection'],
  'Ready for Collection': ['Payment Done'],
  'Payment Done': ['Delivered']
};

// Customer-facing order lifecycle. Order status doubles as each timeline
// notification's Type, so the dashboard can rebuild a per-order timeline by
// filtering Notifications on OrderID instead of storing a separate log.
// Negotiation is a loop, not a forward step — a customer rejecting a
// quotation lands here (with their reason, if given) instead of silently
// looking identical to a brand-new never-quoted order; saveQuotation can be
// called from here just like from Quotation Requested, sending a revised
// quote and moving the order back to Quotation Sent.
var ORDER_STATUSES = [
  'Quotation Requested', 'Quotation Sent', 'Negotiation', 'Order Confirmed',
  'Measurement Taken', 'Work In Progress', 'Ready for Collection',
  'Payment Done', 'Delivered', 'Cancelled'
];

var ACTIONS = {
  createUser: createUser,
  loginUser: loginUser,
  loginAdmin: loginAdmin,
  createOrder: createOrder,
  getOrdersByUser: getOrdersByUser,
  getAllOrders: getAllOrders,
  updateOrderStatus: updateOrderStatus,
  saveQuotation: saveQuotation,
  respondToQuotation: respondToQuotation,
  getQuotationsByUser: getQuotationsByUser,
  getQuotationsByOrder: getQuotationsByOrder,
  saveMeasurements: saveMeasurements,
  getMeasurementsByUser: getMeasurementsByUser,
  getNotifications: getNotifications,
  markNotificationRead: markNotificationRead,
  createWalkInProfile: createWalkInProfile,
  getAllUsers: getAllUsers,
  getOrderTimeline: getOrderTimeline,
  updateUserProfile: updateUserProfile,
  changePassword: changePassword,
  resetUserPassword: resetUserPassword,
  createClass: createClass,
  updateClass: updateClass,
  getPublishedClasses: getPublishedClasses,
  getAllClasses: getAllClasses,
  createEnrollment: createEnrollment,
  getAllEnrollments: getAllEnrollments,
  updateEnrollmentStatus: updateEnrollmentStatus,
  deleteClass: deleteClass,
  deleteCustomer: deleteCustomer,
  recordAdvancePayment: recordAdvancePayment
};

/** One-time setup: creates every tab from SHEET_HEADERS if it doesn't already exist. */
function setupSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
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

function getAllUsers(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var users = sheetToObjects_(getSheet_(SHEET_NAMES.USERS)).map(omitCredentials_);
  return ok_({ users: users });
}

function omitCredentials_(user) {
  var copy = {};
  Object.keys(user).forEach(function (key) {
    if (key !== 'PasswordHash' && key !== 'Salt') copy[key] = user[key];
  });
  return copy;
}

function loginAdmin(data) {
  var email = (data.email || '').trim().toLowerCase();
  var password = data.password || '';

  if (!email || !password) {
    return fail_('Email and password are required');
  }

  var admin = findAdminByEmail_(email);
  if (!admin || hashPassword_(password, admin.Salt) !== admin.PasswordHash) {
    return fail_('Incorrect email or password');
  }

  return ok_({ adminId: admin.AdminID, name: admin.Name, email: admin.Email });
}

function updateUserProfile(data) {
  var userId = (data.userId || '').trim();
  var name = (data.name || '').trim();
  var mobile = (data.mobile || '').trim();
  var city = (data.city || '').trim();

  if (!userId || !name || !mobile) {
    return fail_('Name and mobile are required');
  }

  var found = findById_(SHEET_NAMES.USERS, 'UserID', userId);
  if (!found) {
    return fail_('User not found');
  }

  var existingPhone = findUserByPhone_(mobile);
  if (existingPhone && existingPhone.UserID !== userId) {
    return fail_('Another account already uses this mobile number');
  }

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.USERS);
    var headers = SHEET_HEADERS.Users;
    sheet.getRange(found.rowIndex, headers.indexOf('Name') + 1).setValue(name);
    sheet.getRange(found.rowIndex, headers.indexOf('Phone') + 1).setValue(mobile);
    sheet.getRange(found.rowIndex, headers.indexOf('City') + 1).setValue(city);
  });

  return ok_({ userId: userId, name: name, mobile: mobile, city: city });
}

function changePassword(data) {
  var userId = (data.userId || '').trim();
  var currentPassword = data.currentPassword || '';
  var newPassword = data.newPassword || '';

  if (!userId || !currentPassword || !newPassword) {
    return fail_('Current password and new password are required');
  }
  if (newPassword.length < 8) {
    return fail_('New password must be at least 8 characters');
  }

  var found = findById_(SHEET_NAMES.USERS, 'UserID', userId);
  if (!found) {
    return fail_('User not found');
  }
  if (hashPassword_(currentPassword, found.row.Salt) !== found.row.PasswordHash) {
    return fail_('Current password is incorrect');
  }

  var newSalt = makeSalt_();
  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.USERS);
    var headers = SHEET_HEADERS.Users;
    sheet.getRange(found.rowIndex, headers.indexOf('PasswordHash') + 1).setValue(hashPassword_(newPassword, newSalt));
    sheet.getRange(found.rowIndex, headers.indexOf('Salt') + 1).setValue(newSalt);
  });

  return ok_({ userId: userId });
}

// ---------- Orders ----------

function createOrder(data) {
  var userId = (data.userId || '').trim();
  var garment = (data.service || '').trim();

  if (!userId || !garment) {
    return fail_('userId and service are required');
  }
  if (!findById_(SHEET_NAMES.USERS, 'UserID', userId)) {
    return fail_('User not found');
  }

  var hasMeasurements = parseBool_(data.hasMeasurements);
  var measurementsObj = parseMaybeJson_(data.measurements);
  if (hasMeasurements && measurementsObj && Object.keys(measurementsObj).length) {
    saveMeasurementsRecord_(userId, garment, measurementsObj);
  }

  // Status is always server-set to 'Quotation Requested' on creation — a
  // client-supplied status is never trusted, even though the booking form
  // sends one for clarity.
  var order = createOrderRecord_({
    userId: userId,
    garment: garment,
    notes: data.specialInstructions,
    date: data.preferredDate,
    hasMeasurements: hasMeasurements
  });

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
  var photos = parseMaybeJson_(data.photos); // array of "data:image/...;base64,..." strings, only for Ready for Collection

  if (!orderId || !status) {
    return fail_('orderId and status are required');
  }
  if (ORDER_STATUSES.indexOf(status) === -1) {
    return fail_('Invalid status. Must be one of: ' + ORDER_STATUSES.join(', '));
  }

  var existing = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!existing) {
    return fail_('Order not found');
  }
  if (!isValidStatusTransition_(existing.row.Status, status)) {
    return fail_('Cannot move an order from "' + existing.row.Status + '" to "' + status + '"');
  }

  if (status === 'Ready for Collection') {
    if (!photos || !photos.length) {
      return fail_('At least one photo is required when marking an order Ready for Collection');
    }
    if (photos.length > 3) {
      return fail_('You can upload up to 3 photos');
    }
  }

  var finalAmount = 0;
  if (status === 'Payment Done') {
    finalAmount = Number(data.finalAmount);
    if (!finalAmount || finalAmount < 0) {
      return fail_('finalAmount is required to mark payment done');
    }
  }

  var order = setOrderStatus_(orderId, status);

  var warning = '';
  if (status === 'Ready for Collection' && photos && photos.length) {
    try {
      setOrderPhotoUrls_(orderId, uploadOrderPhotos_(orderId, photos));
    } catch (err) {
      warning = 'Status updated, but photo upload failed: ' + (err && err.message ? err.message : String(err));
    }
  }

  var invoiceUrl = '';
  if (status === 'Payment Done') {
    try {
      invoiceUrl = generateInvoice_(orderId, finalAmount);
      setOrderPaymentDone_(orderId, finalAmount, invoiceUrl);
    } catch (err) {
      warning = 'Status updated, but invoice generation failed: ' + (err && err.message ? err.message : String(err));
    }
  }

  createNotification_(order.UserID, 'Your order status has been updated to "' + status + '".', status, orderId);
  return warning ? fail_(warning) : ok_({ orderId: orderId, status: status, invoiceUrl: invoiceUrl });
}

function recordAdvancePayment(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var orderId = (data.orderId || '').trim();
  var amount = Number(data.amount);
  if (!orderId || !amount || amount <= 0) {
    return fail_('orderId and a positive amount are required');
  }

  var found = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!found) {
    return fail_('Order not found');
  }

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.ORDERS);
    var headers = SHEET_HEADERS.Orders;
    sheet.getRange(found.rowIndex, headers.indexOf('AdvanceAmount') + 1).setValue(amount);
    sheet.getRange(found.rowIndex, headers.indexOf('AdvancePaidAt') + 1).setValue(new Date());
  });

  createNotification_(found.row.UserID, 'We\'ve recorded your advance payment of ₹' + amount + '.', 'Advance Payment Recorded', orderId);
  return ok_({ orderId: orderId, amount: amount });
}

function setOrderPaymentDone_(orderId, finalAmount, invoiceUrl) {
  var found = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!found) return;

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.ORDERS);
    var headers = SHEET_HEADERS.Orders;
    sheet.getRange(found.rowIndex, headers.indexOf('FinalAmountPaid') + 1).setValue(finalAmount);
    sheet.getRange(found.rowIndex, headers.indexOf('PaymentDoneAt') + 1).setValue(new Date());
    sheet.getRange(found.rowIndex, headers.indexOf('InvoiceURL') + 1).setValue(invoiceUrl);
  });
}

/**
 * Builds a PDF invoice (Google Doc -> PDF export, same Drive-folder pattern
 * as order photos) and returns its shareable view URL. There's no WhatsApp
 * Business API in this project, so nothing can auto-attach this file to a
 * WhatsApp message the way a person manually sending a PDF can — the
 * frontend instead sends a wa.me message containing this URL as a link the
 * customer taps to download, and shows the same link as a "Download
 * Invoice" button on their dashboard.
 */
function generateInvoice_(orderId, finalAmount) {
  var orderFound = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!orderFound) throw new Error('Order not found');
  var order = orderFound.row;

  var userFound = findById_(SHEET_NAMES.USERS, 'UserID', order.UserID);
  var user = userFound ? userFound.row : { Name: '', Phone: '', Email: '' };

  var quotation = sheetToObjects_(getSheet_(SHEET_NAMES.QUOTATIONS))
    .filter(function (q) { return q.OrderID === orderId && q.Status === 'Approved'; })
    .sort(byCreatedAtDesc_)[0] || null;

  var advance = Number(order.AdvanceAmount) || 0;
  var tz = Session.getScriptTimeZone();

  var doc = DocumentApp.create('Invoice - ' + orderId);
  var body = doc.getBody();
  body.appendParagraph('Manasa Tailor').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('Home Boutique — Tumkur, Karnataka').setFontSize(10);
  body.appendParagraph('');
  body.appendParagraph('Invoice for Order ' + orderId);
  body.appendParagraph('Date: ' + Utilities.formatDate(new Date(), tz, 'dd MMM yyyy'));
  body.appendParagraph('');
  body.appendParagraph('Bill To: ' + (user.Name || ''));
  body.appendParagraph('Mobile: ' + (user.Phone || ''));
  body.appendParagraph('Service: ' + (order.Garment || ''));
  body.appendParagraph('');

  var rows = [['Item', 'Amount (₹)']];
  if (quotation) {
    rows.push(['Stitching Charges', String(quotation.StitchingCost || 0)]);
    rows.push(['Fabric Cost', String(quotation.FabricCost || 0)]);
    if (Number(quotation.AdditionalCost) > 0) rows.push(['Additional Charges', String(quotation.AdditionalCost)]);
    rows.push(['Total Quoted', String(quotation.TotalAmount || 0)]);
  }
  if (advance > 0) rows.push(['Advance Paid', String(advance)]);
  rows.push(['Amount Paid Now', String(finalAmount)]);
  body.appendTable(rows);

  body.appendParagraph('');
  body.appendParagraph('Thank you for choosing Manasa Tailor!').setItalic(true);

  doc.saveAndClose();
  var pdfBlob = doc.getAs('application/pdf');
  var docFile = DriveApp.getFileById(doc.getId());

  var folder = getInvoicesFolder_();
  var pdfFile = folder.createFile(pdfBlob).setName('Invoice-' + orderId + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  docFile.setTrashed(true); // only the exported PDF is kept, not the source Doc

  return 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId();
}

function getInvoicesFolder_() {
  var name = 'Manasa Tailor Invoices';
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function isValidStatusTransition_(fromStatus, toStatus) {
  if (toStatus === 'Cancelled') {
    return fromStatus !== 'Delivered' && fromStatus !== 'Cancelled';
  }
  var allowed = VALID_STATUS_TRANSITIONS[fromStatus];
  return !!allowed && allowed.indexOf(toStatus) !== -1;
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

/**
 * Order photos are uploaded as base64 and stored in Drive, not the Sheet —
 * a single Sheet cell caps out around 50,000 characters, which even one
 * compressed phone photo's base64 easily exceeds. Orders.PhotoURLs just
 * holds the resulting comma-separated Drive view URLs, matching what it
 * already expected from earlier (manually-pasted) usage.
 */
function uploadOrderPhotos_(orderId, photos) {
  var folder = getOrderPhotosFolder_();
  return photos.map(function (photo, i) {
    var blob = base64ImageToBlob_(photo, 'order-' + orderId + '-' + (i + 1));
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  });
}

function getOrderPhotosFolder_() {
  var name = 'Manasa Tailor Order Photos';
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function base64ImageToBlob_(dataUri, filenameBase) {
  var match = /^data:(image\/(png|jpe?g));base64,(.+)$/i.exec(dataUri || '');
  if (!match) {
    throw new Error('Photo must be a base64 PNG or JPEG data URL');
  }
  var mimeType = match[1];
  var ext = mimeType.toLowerCase().indexOf('png') !== -1 ? 'png' : 'jpg';
  var bytes = Utilities.base64Decode(match[3]);
  return Utilities.newBlob(bytes, mimeType, filenameBase + '.' + ext);
}

function setOrderPhotoUrls_(orderId, newUrls) {
  var found = findById_(SHEET_NAMES.ORDERS, 'OrderID', orderId);
  if (!found) return;

  var existing = String(found.row.PhotoURLs || '').split(',').map(function (u) { return u.trim(); }).filter(Boolean);
  var combined = existing.concat(newUrls).slice(0, 3);

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.ORDERS);
    var headers = SHEET_HEADERS.Orders;
    sheet.getRange(found.rowIndex, headers.indexOf('PhotoURLs') + 1).setValue(combined.join(','));
  });
}

// ---------- Quotations ----------

function saveQuotation(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

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
  var requiredAdvance = Number(data.requiredAdvance) || 0;
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
    DeliveryTimeline: (data.deliveryTimeline || '').trim(),
    AdminNotes: (data.adminNotes || '').trim(),
    RequiredAdvance: requiredAdvance,
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
  var reason = (data.reason || '').trim();

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
  // Rejecting lands on "Negotiation" rather than back on "Quotation
  // Requested" — otherwise a rejected quote is indistinguishable from a
  // brand-new never-quoted order in the admin dashboard, with no record of
  // what the customer actually wanted changed.
  var newOrderStatus = decision === 'approve' ? 'Order Confirmed' : 'Negotiation';

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.QUOTATIONS);
    var headers = SHEET_HEADERS.Quotations;
    sheet.getRange(quotation.rowIndex, headers.indexOf('Status') + 1).setValue(quotationStatus);
    sheet.getRange(quotation.rowIndex, headers.indexOf('UpdatedAt') + 1).setValue(new Date());
  });

  setOrderStatus_(orderId, newOrderStatus);

  if (decision === 'reject') {
    withLock_(function () {
      var sheet = getSheet_(SHEET_NAMES.ORDERS);
      var headers = SHEET_HEADERS.Orders;
      sheet.getRange(order.rowIndex, headers.indexOf('NegotiationNote') + 1).setValue(reason);
    });
  }

  createNotification_(
    userId,
    decision === 'approve'
      ? 'You approved the quotation — your order is now confirmed.'
      : 'You requested changes to the quotation. We\'ll be in touch to discuss.',
    newOrderStatus,
    orderId
  );
  notifyAdmin_(
    decision === 'approve' ? 'Customer Approved' : 'Negotiation Requested',
    decision === 'approve'
      ? 'A customer approved the quotation for order ' + orderId + '.'
      : 'A customer requested changes for order ' + orderId + (reason ? ': "' + reason + '"' : '.'),
    orderId
  );

  return ok_({ orderId: orderId, quotationId: quotationId, status: newOrderStatus });
}

function getQuotationsByUser(data) {
  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }

  // AdminNotes is internal-only — strip it before this customer-facing
  // response leaves the server, not just hide it in the UI.
  var quotations = sheetToObjects_(getSheet_(SHEET_NAMES.QUOTATIONS))
    .filter(function (row) { return row.UserID === userId; })
    .sort(byCreatedAtDesc_)
    .map(omitAdminNotes_);

  return ok_({ quotations: quotations });
}

function omitAdminNotes_(quotation) {
  var copy = {};
  Object.keys(quotation).forEach(function (key) {
    if (key !== 'AdminNotes') copy[key] = quotation[key];
  });
  return copy;
}

/**
 * Notifications are tagged with OrderID regardless of who they were sent to,
 * so this returns the full mixed audit trail (customer status updates AND
 * admin-only events like "New Quotation Request") for one order — useful
 * for admin.html's timeline, since a customer's own getNotifications only
 * ever sees rows addressed to their own userId, never another order event.
 */
function getOrderTimeline(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var orderId = (data.orderId || '').trim();
  if (!orderId) {
    return fail_('orderId is required');
  }

  var events = sheetToObjects_(getSheet_(SHEET_NAMES.NOTIFICATIONS))
    .filter(function (row) { return row.OrderID === orderId; })
    .sort(byCreatedAtDesc_);

  return ok_({ events: events });
}

function getQuotationsByOrder(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var orderId = (data.orderId || '').trim();
  if (!orderId) {
    return fail_('orderId is required');
  }

  var quotations = sheetToObjects_(getSheet_(SHEET_NAMES.QUOTATIONS))
    .filter(function (row) { return row.OrderID === orderId; })
    .sort(byCreatedAtDesc_);

  return ok_({ quotations: quotations });
}

// ---------- Measurements ----------

function saveMeasurements(data) {
  var userId = (data.userId || '').trim();
  var garment = (data.garment || '').trim();
  var measurements = parseMaybeJson_(data.measurements);

  if (!userId || !garment || !measurements || !Object.keys(measurements).length) {
    return fail_('userId, garment and a measurements object are required');
  }
  if (!findById_(SHEET_NAMES.USERS, 'UserID', userId)) {
    return fail_('User not found');
  }

  saveMeasurementsRecord_(userId, garment, measurements);
  return ok_({ saved: true });
}

/** Upserts a Measurements row for (userId, garment). Shared by saveMeasurements and createOrder. */
function saveMeasurementsRecord_(userId, garment, measurements) {
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

  if (parseBool_(data.markAllRead)) {
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

  var unreadOnly = parseBool_(data.unreadOnly);
  var notifications = rows
    .filter(function (row) {
      if (row.UserID !== userId) return false;
      if (unreadOnly) return !row.IsRead;
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

// ---------- Classes ----------

/**
 * TimingSlots (an array like ['Morning','Weekend'] from the admin panel's
 * checkboxes) arrives the same way every other array/object param does over
 * this app's GET transport — base64-encoded JSON, see parseMaybeJson_'s own
 * comment for why raw JSON in a query string isn't safe on this deployment.
 * Stored in the sheet as a plain comma-joined string rather than another
 * JSON blob column, since it's just a handful of fixed values and a human
 * glancing at the Classes tab shouldn't have to mentally parse JSON to read
 * it.
 */
function formatTimingSlots_(value) {
  var arr = parseMaybeJson_(value);
  if (arr && Array.isArray(arr)) return arr.join(',');
  return typeof value === 'string' ? value.trim() : '';
}

function createClass(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var title = (data.title || '').trim();
  var level = (data.level || '').trim();
  if (!title || !level) {
    return fail_('title and level are required');
  }

  var status = (data.status || 'Draft').trim();
  if (CLASS_STATUSES.indexOf(status) === -1) {
    return fail_('Invalid status. Must be one of: ' + CLASS_STATUSES.join(', '));
  }

  var now = new Date();
  var record = {
    ClassID: genId_('CLS'),
    Title: title,
    Level: level,
    Description: (data.description || '').trim(),
    Duration: (data.duration || '').trim(),
    TimingSlots: formatTimingSlots_(data.timingSlots),
    StartDate: data.startDate || '',
    TotalSeats: Number(data.totalSeats) || 0,
    SeatsTaken: Number(data.seatsTaken) || 0,
    ContactNote: (data.contactNote || 'Call to enquire about fees').trim(),
    Fees: Number(data.fees) || 0,
    Status: status,
    CreatedAt: now,
    UpdatedAt: now
  };

  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.CLASSES), SHEET_HEADERS.Classes, record);
  });

  if (status === 'Published') {
    notifyAllCustomersOfNewClass_(title);
  }

  return ok_({ classId: record.ClassID });
}

/**
 * Every existing customer gets a Notification row when a class newly
 * becomes Published — a single batch write (one setValues call covering
 * every row) rather than one appendRow_/withLock_ per customer, since a
 * lock-per-row loop would serialize badly once there are more than a
 * handful of accounts. Deliberately not sent on every edit of an
 * already-published class — see the "previousStatus" guard at both call
 * sites — so tweaking a typo doesn't re-spam everyone.
 */
function notifyAllCustomersOfNewClass_(classTitle) {
  var users = sheetToObjects_(getSheet_(SHEET_NAMES.USERS));
  if (!users.length) return;

  var now = new Date();
  var headers = SHEET_HEADERS.Notifications;
  var message = 'New tailoring class now open: "' + classTitle + '". Check it out!';
  var rows = users.map(function (u) {
    var record = {
      NotificationID: genId_('NOT'),
      UserID: u.UserID,
      Message: message,
      Type: 'Class Published',
      IsRead: false,
      CreatedAt: now,
      OrderID: ''
    };
    return headers.map(function (h) { return record.hasOwnProperty(h) ? record[h] : ''; });
  });

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.NOTIFICATIONS);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  });
}

/**
 * Partial update, deliberately — admin.html's class cards have quick
 * publish/unpublish/seats-taken actions that shouldn't have to resend every
 * field just to flip one, and a naive full-overwrite would blank out
 * whatever a quick action's payload happened to omit.
 */
function updateClass(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var classId = (data.classId || '').trim();
  if (!classId) {
    return fail_('classId is required');
  }
  var found = findById_(SHEET_NAMES.CLASSES, 'ClassID', classId);
  if (!found) {
    return fail_('Class not found');
  }

  if (data.status !== undefined && CLASS_STATUSES.indexOf((data.status || '').trim()) === -1) {
    return fail_('Invalid status. Must be one of: ' + CLASS_STATUSES.join(', '));
  }

  var fieldSetters = {
    title: function (v) { return { Title: (v || '').trim() }; },
    level: function (v) { return { Level: (v || '').trim() }; },
    description: function (v) { return { Description: (v || '').trim() }; },
    duration: function (v) { return { Duration: (v || '').trim() }; },
    timingSlots: function (v) { return { TimingSlots: formatTimingSlots_(v) }; },
    startDate: function (v) { return { StartDate: v || '' }; },
    totalSeats: function (v) { return { TotalSeats: Number(v) || 0 }; },
    seatsTaken: function (v) { return { SeatsTaken: Number(v) || 0 }; },
    contactNote: function (v) { return { ContactNote: (v || '').trim() }; },
    fees: function (v) { return { Fees: Number(v) || 0 }; },
    status: function (v) { return { Status: (v || '').trim() }; }
  };

  var updates = {};
  Object.keys(fieldSetters).forEach(function (key) {
    if (data[key] === undefined || data[key] === null) return;
    var partial = fieldSetters[key](data[key]);
    Object.keys(partial).forEach(function (col) { updates[col] = partial[col]; });
  });
  updates.UpdatedAt = new Date();

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.CLASSES);
    var headers = SHEET_HEADERS.Classes;
    Object.keys(updates).forEach(function (col) {
      sheet.getRange(found.rowIndex, headers.indexOf(col) + 1).setValue(updates[col]);
    });
  });

  if (updates.Status === 'Published' && found.row.Status !== 'Published') {
    notifyAllCustomersOfNewClass_(updates.Title || found.row.Title);
  }

  return ok_({ classId: classId });
}

/** Public — this is what classes.html and index.html's teaser both fetch, logged in or not. */
function getPublishedClasses(data) {
  var classes = sheetToObjects_(getSheet_(SHEET_NAMES.CLASSES))
    .filter(function (row) { return row.Status === 'Published'; })
    .sort(byCreatedAtDesc_);
  return ok_({ classes: classes });
}

function getAllClasses(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var classes = sheetToObjects_(getSheet_(SHEET_NAMES.CLASSES)).sort(byCreatedAtDesc_);
  return ok_({ classes: classes });
}

// ---------- Class enrollments ----------

function createEnrollment(data) {
  var classId = (data.classId || '').trim();
  var userId = (data.userId || '').trim();
  var customerName = (data.customerName || '').trim();
  var mobile = (data.mobile || '').trim();

  if (!classId || !userId || !customerName || !mobile) {
    return fail_('classId, userId, customerName and mobile are required');
  }
  if (!findById_(SHEET_NAMES.CLASSES, 'ClassID', classId)) {
    return fail_('Class not found');
  }
  if (!findById_(SHEET_NAMES.USERS, 'UserID', userId)) {
    return fail_('User not found');
  }

  var record = {
    EnrollmentID: genId_('ENR'),
    ClassID: classId,
    UserID: userId,
    CustomerName: customerName,
    Mobile: mobile,
    Email: (data.email || '').trim(),
    PreferredLevel: (data.preferredLevel || '').trim(),
    PreferredTiming: (data.preferredTiming || '').trim(),
    Message: (data.message || '').trim(),
    Status: 'New',
    AdminNotes: '',
    CreatedAt: new Date()
  };

  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.CLASS_ENROLLMENTS), SHEET_HEADERS.ClassEnrollments, record);
  });

  notifyAdmin_('New Class Enrollment', customerName + ' is interested in a tailoring class.', '');

  return ok_({ enrollmentId: record.EnrollmentID });
}

function getAllEnrollments(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var enrollments = sheetToObjects_(getSheet_(SHEET_NAMES.CLASS_ENROLLMENTS)).sort(byCreatedAtDesc_);
  return ok_({ enrollments: enrollments });
}

function updateEnrollmentStatus(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var enrollmentId = (data.enrollmentId || '').trim();
  var status = (data.status || '').trim();
  if (!enrollmentId || !status) {
    return fail_('enrollmentId and status are required');
  }
  if (ENROLLMENT_STATUSES.indexOf(status) === -1) {
    return fail_('Invalid status. Must be one of: ' + ENROLLMENT_STATUSES.join(', '));
  }
  var found = findById_(SHEET_NAMES.CLASS_ENROLLMENTS, 'EnrollmentID', enrollmentId);
  if (!found) {
    return fail_('Enrollment not found');
  }

  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.CLASS_ENROLLMENTS);
    var headers = SHEET_HEADERS.ClassEnrollments;
    sheet.getRange(found.rowIndex, headers.indexOf('Status') + 1).setValue(status);
    if (data.notes !== undefined && data.notes !== null) {
      sheet.getRange(found.rowIndex, headers.indexOf('AdminNotes') + 1).setValue((data.notes || '').trim());
    }
  });

  return ok_({ enrollmentId: enrollmentId, status: status });
}

/**
 * Not in the original 7-function spec for this feature, but admin.html's
 * class cards need a real "Delete" action (Draft classes especially are
 * scratch entries an admin should be able to fully remove, not just hide).
 * Matches this app's existing bias against silent data loss elsewhere
 * (Orders get Cancelled rather than deleted, Users are never deleted) by
 * refusing to delete a class that already has enrollment history — that's
 * what Closed is for instead.
 */
function deleteClass(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var classId = (data.classId || '').trim();
  if (!classId) {
    return fail_('classId is required');
  }
  var found = findById_(SHEET_NAMES.CLASSES, 'ClassID', classId);
  if (!found) {
    return fail_('Class not found');
  }

  var hasEnrollments = sheetToObjects_(getSheet_(SHEET_NAMES.CLASS_ENROLLMENTS))
    .some(function (row) { return row.ClassID === classId; });
  if (hasEnrollments) {
    return fail_('This class has enrollments and can\'t be deleted — set it to Closed instead.');
  }

  withLock_(function () {
    getSheet_(SHEET_NAMES.CLASSES).deleteRow(found.rowIndex);
  });

  return ok_({ classId: classId });
}

// ---------- Shared record creation ----------

function createUserRecord_(name, phone, email, password, city, username) {
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
    City: city || '',
    Username: username || ''
  };
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.USERS), SHEET_HEADERS.Users, record);
  });
  return record;
}

function createOrderRecord_(opts) {
  var now = new Date();
  var record = {
    OrderID: genId_('ORD'),
    UserID: opts.userId,
    Garment: opts.garment,
    Purpose: opts.purpose || '',
    Date: opts.date || '',
    TimeSlot: opts.slot || '',
    Notes: opts.notes || '',
    Status: 'Quotation Requested',
    CreatedAt: now,
    UpdatedAt: now,
    PhotoURLs: '',
    HasMeasurements: opts.hasMeasurements ? 'Yes' : 'No'
  };
  withLock_(function () {
    appendRow_(getSheet_(SHEET_NAMES.ORDERS), SHEET_HEADERS.Orders, record);
  });
  createNotification_(opts.userId, 'Your order for ' + opts.garment + ' has been received.', 'Quotation Requested', record.OrderID);
  notifyAdmin_('New Quotation Request', 'A new quotation request came in for ' + opts.garment + '.', record.OrderID);
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

/** Notifies the boutique's one admin account (if one has been set up yet). */
function notifyAdmin_(type, message, orderId) {
  var adminId = getPrimaryAdminId_();
  if (adminId) {
    createNotification_(adminId, message, type, orderId);
  }
}

// ---------- Admin ----------

function requireAdmin_(adminId) {
  adminId = (adminId || '').trim();
  if (!adminId || !findById_(SHEET_NAMES.ADMIN, 'AdminID', adminId)) {
    return fail_('Unauthorized: valid adminId required');
  }
  return null;
}

function findAdminByEmail_(email) {
  var rows = sheetToObjects_(getSheet_(SHEET_NAMES.ADMIN));
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].Email && rows[i].Email.toLowerCase() === email) return rows[i];
  }
  return null;
}

/** The Admin sheet is expected to hold a single row (see createAdminAccount()). */
function getPrimaryAdminId_() {
  var rows = sheetToObjects_(getSheet_(SHEET_NAMES.ADMIN));
  return rows.length ? rows[0].AdminID : null;
}

function createWalkInProfile(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var name = (data.name || '').trim();
  var phone = (data.mobile || '').trim();
  var email = (data.email || '').trim().toLowerCase();
  var reason = (data.reason || '').trim();
  var service = (data.service || '').trim();

  if (!name || !phone || !email) {
    return fail_('Name, mobile and email are required');
  }
  if (findUserByPhone_(phone)) {
    return fail_('A profile with this mobile number already exists');
  }
  if (findUserByEmail_(email)) {
    return fail_('A profile with this email already exists');
  }

  var username = generateWalkInUsername_(name, phone);
  var password = generateWalkInPassword_();
  var user = createUserRecord_(name, phone, email, password, '', username);

  var orderId = null;
  if (service) {
    orderId = createOrderRecord_({ userId: user.UserID, garment: service, purpose: reason, notes: reason }).OrderID;
  }

  // The generated "Username" is a friendly label for the credentials slip —
  // actual login (login.html) is email + password, same as self-registered
  // customers, so the account works through the one login path that exists
  // rather than needing a second, username-based auth system.
  return ok_({
    userId: user.UserID,
    username: username,
    password: password,
    email: email,
    orderId: orderId
  });
}

/**
 * Admin-issued password reset. Generates a fresh random password the same
 * way createWalkInProfile does, so there's always a way to get a customer
 * a working password again — including the walk-in case where the
 * original generated password was shown only once and never stored in
 * plain text, so a lost response there is otherwise unrecoverable.
 */
function resetUserPassword(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }

  var found = findById_(SHEET_NAMES.USERS, 'UserID', userId);
  if (!found) {
    return fail_('User not found');
  }

  var newPassword = generateWalkInPassword_();
  var newSalt = makeSalt_();
  withLock_(function () {
    var sheet = getSheet_(SHEET_NAMES.USERS);
    var headers = SHEET_HEADERS.Users;
    sheet.getRange(found.rowIndex, headers.indexOf('PasswordHash') + 1).setValue(hashPassword_(newPassword, newSalt));
    sheet.getRange(found.rowIndex, headers.indexOf('Salt') + 1).setValue(newSalt);
  });

  return ok_({ userId: userId, name: found.row.Name, email: found.row.Email, phone: found.row.Phone, password: newPassword });
}

/**
 * Same "never silently destroy data with dependents" bias as deleteClass:
 * a customer with any order history keeps their account, since deleting it
 * would leave that history's UserID pointing at nothing. There's no
 * "Closed"-style escape hatch for customers the way there is for classes —
 * a customer with orders just isn't deletable from here, full stop.
 */
function deleteCustomer(data) {
  var adminErr = requireAdmin_(data.adminId);
  if (adminErr) return adminErr;

  var userId = (data.userId || '').trim();
  if (!userId) {
    return fail_('userId is required');
  }
  var found = findById_(SHEET_NAMES.USERS, 'UserID', userId);
  if (!found) {
    return fail_('Customer not found');
  }

  var hasOrders = sheetToObjects_(getSheet_(SHEET_NAMES.ORDERS))
    .some(function (row) { return row.UserID === userId; });
  if (hasOrders) {
    return fail_('This customer has order history and can\'t be deleted.');
  }

  withLock_(function () {
    getSheet_(SHEET_NAMES.USERS).deleteRow(found.rowIndex);
  });

  return ok_({ userId: userId });
}

function generateWalkInUsername_(name, phone) {
  var firstName = (name.trim().split(/\s+/)[0] || 'guest').toLowerCase().replace(/[^a-z0-9]/g, '') || 'guest';
  var digits = phone.replace(/\D/g, '');
  var last4 = digits.length >= 4 ? digits.slice(-4) : ('0000' + digits).slice(-4);
  return firstName + '_' + last4;
}

function generateWalkInPassword_() {
  var digits = '';
  for (var i = 0; i < 5; i++) digits += Math.floor(Math.random() * 10);
  return 'MT@' + digits;
}

/**
 * One-time admin bootstrap — NOT in the ACTIONS map, so it's never reachable
 * over the web. Set ADMIN_BOOTSTRAP_EMAIL/NAME/PASSWORD below, run this once
 * from the Apps Script editor (select it in the toolbar dropdown, click Run),
 * then blank out ADMIN_BOOTSTRAP_PASSWORD and re-save. Re-running it after
 * that updates the same admin's password rather than creating a duplicate row.
 */
var ADMIN_BOOTSTRAP_NAME = 'Shilpa';
var ADMIN_BOOTSTRAP_EMAIL = '';
var ADMIN_BOOTSTRAP_PASSWORD = '';

function createAdminAccount() {
  if (!ADMIN_BOOTSTRAP_EMAIL || !ADMIN_BOOTSTRAP_PASSWORD) {
    throw new Error('Set ADMIN_BOOTSTRAP_EMAIL and ADMIN_BOOTSTRAP_PASSWORD at the top of this function before running it.');
  }

  var email = ADMIN_BOOTSTRAP_EMAIL.trim().toLowerCase();
  var salt = makeSalt_();
  var passwordHash = hashPassword_(ADMIN_BOOTSTRAP_PASSWORD, salt);
  var existing = findAdminByEmail_(email);
  var sheet = getSheet_(SHEET_NAMES.ADMIN);
  var headers = SHEET_HEADERS.Admin;

  if (existing) {
    var found = findById_(SHEET_NAMES.ADMIN, 'AdminID', existing.AdminID);
    withLock_(function () {
      sheet.getRange(found.rowIndex, headers.indexOf('PasswordHash') + 1).setValue(passwordHash);
      sheet.getRange(found.rowIndex, headers.indexOf('Salt') + 1).setValue(salt);
    });
    Logger.log('Updated password for existing admin: ' + email);
    return;
  }

  var record = {
    AdminID: genId_('ADM'),
    Name: ADMIN_BOOTSTRAP_NAME || 'Admin',
    Phone: '',
    Email: email,
    PasswordHash: passwordHash,
    Salt: salt,
    CreatedAt: new Date()
  };
  withLock_(function () {
    appendRow_(sheet, headers, record);
  });
  Logger.log('Created admin account: ' + email);
}

// ---------- Sheet helpers ----------

function getSheet_(name) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
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

/**
 * GET query params always arrive as strings (doGet -> e.parameter), unlike
 * POST's JSON.parse(e.postData.contents) which preserves real types — so a
 * boolean param sent over GET shows up here as the *string* "false", which
 * is still truthy in a naive `!!data.x` check. Use this instead anywhere a
 * param is expected to carry a real boolean.
 */
function parseBool_(value) {
  return value === true || value === 'true';
}

/**
 * Same story for object/array params (measurements, photos): over GET
 * they're JSON-stringified into the query string by the client and arrive
 * here as a string that needs parsing back out, whereas POST could already
 * hand over a real object. Returns null if value isn't already an
 * object/array and doesn't parse as one.
 *
 * The string case arrives base64-encoded, not raw JSON — confirmed live
 * that a raw JSON string (braces/quotes/colons/spaces) placed straight into
 * a GET query string doesn't reliably survive this Web App's own parameter
 * parsing (e.g. saveMeasurements with measurements={"Bust":"12 in"} failed
 * validation as if the param were empty, even when pasting that exact URL
 * directly into a browser — no client JS/CORS involved at all). Base64
 * avoids that: its alphabet is just letters, digits, +, /, = — nothing a
 * query-string layer could reinterpret. See callApi/utf8ToBase64_ in
 * config.js for the encode side.
 */
function parseMaybeJson_(value) {
  if (value && typeof value === 'object') return value;
  if (typeof value === 'string' && value.trim()) {
    try {
      var decoded = Utilities.newBlob(Utilities.base64Decode(value)).getDataAsString('UTF-8');
      return JSON.parse(decoded);
    } catch (err) {
      return null;
    }
  }
  return null;
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
