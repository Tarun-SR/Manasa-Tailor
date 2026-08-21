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
  'loginUser', 'loginAdmin', 'markNotificationRead', 'updateUserProfile'
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
