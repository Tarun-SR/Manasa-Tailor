// Manasa Tailor — shared site config. Loaded via <script src="config.js">
// in every page's <head>, before any page-specific script runs, so
// CONFIG.WEBAPP_URL etc. are available wherever they're needed. Update
// values here once rather than in every page that uses them.
const CONFIG = {
  WEBAPP_URL: 'https://script.google.com/macros/s/AKfycbw0TxR0rpAsgeVSOtf70LQW_eE7PTJm2K9DdBXphSt7PqDM9C7Yiy86pbGgNNG2lY5L/exec',
  SITE_URL: 'https://tarun-sr.github.io/Manasa-Tailor/',
  ADMIN_EMAIL: 'manasatailor1982@gmail.com',
  WHATSAPP_NUMBER: '916360194277'
};

/**
 * Shared API helper — GET with query params, not POST.
 *
 * Apps Script Web Apps redirect every request through a
 * script.googleusercontent.com "echo" URL before the actual JSON comes
 * back. In real browser testing that hop reliably works for GET, but not
 * for POST — confirmed live, even though the response headers looked fine
 * under curl (curl only shows headers, it doesn't replicate a browser's
 * full CORS enforcement, so it can't be trusted to validate this).
 *
 * This means every value crosses as a URL string — objects/arrays get
 * JSON.stringify'd here and must be parsed back out server-side
 * (parseMaybeJson_ in apps_script.gs), and booleans arrive as the literal
 * strings "true"/"false" rather than real booleans (parseBool_ handles
 * that server-side too).
 *
 * NOT used for: anything carrying photo data (base64 images blow well past
 * this endpoint's ~8-16KB query-string ceiling — see admin.html's photo
 * upload, which stays on a fire-and-forget POST for that reason) or
 * anywhere a smaller, equally-safe alternative exists. It IS used for
 * login/register/change-password, which means those passwords travel as
 * plain URL query parameters — visible in browser history — because POST
 * doesn't work at all on this deployment. That's a real, known tradeoff,
 * not an oversight.
 */
// Read-only or otherwise side-effect-free actions — safe to silently retry
// once on failure, since running them twice changes nothing. Actions that
// create/mutate data (saveQuotation, createOrder, updateOrderStatus, ...)
// are deliberately left out: confirmed live that a "failed" write can
// still have gone through server-side (the response just didn't make it
// back cleanly), so blindly retrying one of those could create a
// duplicate record instead of just re-reading the same answer. Callers of
// those actions need to verify against a safe read before deciding
// whether to actually retry — see admin.html's sendQuotation for the
// pattern.
var CALLAPI_SAFE_RETRY_ACTIONS = [
  'getOrdersByUser', 'getAllOrders', 'getQuotationsByUser', 'getQuotationsByOrder',
  'getMeasurementsByUser', 'getNotifications', 'getAllUsers', 'getOrderTimeline',
  'loginUser', 'loginAdmin', 'markNotificationRead', 'updateUserProfile',
  // Each call just overwrites the same password field — no duplicate
  // resource risk the way createUser/createOrder/etc. have, so unlike
  // those, this one really is safe to retry blindly. Only the result the
  // admin actually sees (the last call that got a response back) matters.
  'resetUserPassword',
  // Same reasoning: updateOrderStatus overwrites the same Status cell on
  // the same order row, and saveMeasurements upserts by (userId, garment)
  // rather than appending — see setOrderStatus_ / saveMeasurementsRecord_
  // in apps_script.gs. Replaying either with the same input twice leaves
  // the sheet in the same state as replaying it once, so a blind retry
  // here can't create a duplicate the way createOrder/createWalkInProfile
  // could. This is what actually closes the "false failure" gap for the
  // admin status-update flow: without it, a dropped connection on this
  // call had only the caller's own verify-by-read fallback to fall back
  // on, and if that read also hiccupped on this transport, the admin saw
  // a failure for a write that had, in fact, gone through.
  'updateOrderStatus', 'saveMeasurements'
];

function callApi(action, params) {
  params = params || {};
  var url = new URL(CONFIG.WEBAPP_URL);
  url.searchParams.set('action', action);
  Object.keys(params).forEach(function (key) {
    var value = params[key];
    if (value === undefined || value === null) return;
    url.searchParams.set(key, typeof value === 'object' ? JSON.stringify(value) : String(value));
  });

  function attempt() {
    return fetch(url.toString()).then(function (res) {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    });
  }

  if (CALLAPI_SAFE_RETRY_ACTIONS.indexOf(action) === -1) {
    return attempt();
  }
  return attempt().catch(function () {
    return new Promise(function (resolve) { setTimeout(resolve, 700); }).then(attempt);
  });
}

/**
 * Opens the tab a WhatsApp handoff will eventually land in, and immediately
 * paints a "Redirecting…" placeholder into it — call this synchronously
 * inside the click handler (before any await/then), same as the bare
 * window.open('', '_blank') calls it replaces, so the user gesture still
 * carries through to the popup blocker.
 *
 * The pattern across this app is: open the tab first, do some async work
 * (save to the sheet, etc.), then set tab.location.href to the real wa.me
 * link once that work resolves. Left blank in between, that tab just shows
 * about:blank / a blank white page for however long the async step takes —
 * which reads as broken rather than as "give it a second". This fills that
 * gap with a page that names the destination so it's obvious where the
 * user is headed and that something is happening, not stuck.
 */
function openWhatsAppRedirectTab_() {
  var tab = window.open('', '_blank');
  if (tab) {
    tab.document.write(
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<title>Redirecting to WhatsApp…</title>' +
      '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
      '<style>' +
      'body{margin:0;min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;' +
      'gap:16px;padding:24px;text-align:center;font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;' +
      'background:#fdf8f5;color:#2c1a20}' +
      '.spinner{width:36px;height:36px;border:3px solid #e8d5da;border-top-color:#25d366;border-radius:50%;' +
      'animation:wa-spin .8s linear infinite}' +
      '@keyframes wa-spin{to{transform:rotate(360deg)}}' +
      'p{margin:0;font-size:14px;color:#7a5a64;max-width:320px;line-height:1.6}' +
      'strong{color:#2c1a20}' +
      '</style></head><body>' +
      '<div class="spinner" role="status" aria-label="Redirecting"></div>' +
      '<p><strong>Taking you to WhatsApp…</strong><br>This only takes a moment.</p>' +
      '</body></html>'
    );
    tab.document.close();
  }
  return tab;
}
