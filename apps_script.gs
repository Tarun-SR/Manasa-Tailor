/**
 * Manasa Tailor — Order Enquiry intake
 *
 * This script receives POSTs from the "Custom Order Enquiry" form on the
 * website and appends each submission as a new row in a Google Sheet.
 *
 * SETUP:
 * 1. Create a new Google Sheet (e.g. "Manasa Tailor Orders").
 *    Add a header row: Timestamp | Name | Phone | Garment | Purpose | Date | Time Slot | Notes
 * 2. In the Sheet: Extensions -> Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. Click Deploy -> New deployment -> select type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Click Deploy, authorize the permissions Google asks for.
 * 6. Copy the "Web app URL" it gives you.
 * 7. Paste that URL into index.html, replacing:
 *      var GOOGLE_SHEET_WEBAPP_URL = "PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE";
 * 8. Save, commit, and push to GitHub — GitHub Pages will pick it up automatically.
 *
 * Every time you edit this script after that, you must create a NEW
 * deployment (or use "Manage deployments" -> edit -> new version) for the
 * change to take effect — the URL stays the same either way.
 */

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = e.parameter || {};
  }

  sheet.appendRow([
    new Date(),
    data.name || '',
    data.phone || '',
    data.garment || '',
    data.purpose || '',
    data.date || '',
    data.slot || '',
    data.notes || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}