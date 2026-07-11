/**
 * ============================================================
 * MEER HQ — LEADVAULT ORDER BACKEND (Google Apps Script)
 * ============================================================
 *
 * WHAT THIS IS
 * A tiny free backend that lets your static LeadVault page and your
 * new admin.html page share real data — every order lands in a Google
 * Sheet, and admin.html can read that Sheet from anywhere, on any
 * device, not just the browser that placed the order.
 *
 * ============================================================
 * SETUP — DO THIS ONCE (about 5 minutes)
 * ============================================================
 * 1. Go to https://script.google.com → New project.
 * 2. Delete anything in the editor, paste this whole file in.
 * 3. Change SECRET_KEY below to your own long random string
 *    (this is what protects your admin page — treat it like a password).
 * 4. Click "Deploy" → "New deployment".
 * 5. Click the gear icon next to "Select type" → choose "Web app".
 * 6. Fill in:
 *      Execute as:      Me (your Google account)
 *      Who has access:  Anyone
 *    ("Anyone" sounds scary, but the SECRET_KEY below is what actually
 *    protects reading order data — without it, requests are rejected.)
 * 7. Click Deploy. Authorize the permissions it asks for (it's your
 *    own script, this is expected).
 * 8. Copy the "Web app URL" it gives you — looks like:
 *      https://script.google.com/macros/s/AKfycb.../exec
 *    You'll paste this into BOTH leadvault-landing-page.html (as
 *    APPS_SCRIPT_URL) and admin.html (as APPS_SCRIPT_URL + ADMIN_KEY).
 * 9. This will automatically create a new Google Sheet the first time
 *    an order comes in (or the first time you open admin.html) — you
 *    don't need to create one manually. It'll be named "LeadVault Orders"
 *    and land in the Google Drive of the account you deployed with.
 *
 * WHEN YOU UPDATE THIS CODE LATER: after editing, you must click
 * "Deploy" → "Manage deployments" → pencil icon → "New version" →
 * Deploy again, or your changes won't go live. The Web App URL stays
 * the same across redeploys, so you won't need to update the HTML files.
 * ============================================================
 */

// ⚠️ CHANGE THIS — this is what protects your order data.
// Use a long random string, e.g. generate one at https://1password.com/password-generator/
const SECRET_KEY = 'JvKafxy9t0NjipoKTWlblCp7PNWao7IH';

const SHEET_NAME = 'Orders';
const COLUMNS = [
  'Timestamp', 'OrderID', 'Plan', 'Name', 'Email', 'Company', 'Website',
  'Titles', 'Industries', 'Geography', 'CompanySize', 'UseCase', 'Notes',
  'PaymentConfirmed', 'FlwTxRef', 'FlwTransactionId'
];

/**
 * Handles POST requests — used to (a) create a new order, or
 * (b) mark an existing order as paid.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'confirmPayment') {
      markOrderPaid(data.orderId, data.flwTxRef, data.flwTransactionId);
      return jsonResponse({ ok: true, action: 'confirmPayment' });
    }

    // Default action: create a new order row.
    const sheet = getOrCreateSheet();
    sheet.appendRow([
      new Date(),
      data.id || '',
      data.plan || '',
      data.name || '',
      data.email || '',
      data.company || '',
      data.website || '',
      data.titles || '',
      data.industries || '',
      data.geo || '',
      data.companySize || '',
      data.useCase || '',
      data.notes || '',
      'no',   // PaymentConfirmed starts false
      '',     // FlwTxRef
      ''      // FlwTransactionId
    ]);
    return jsonResponse({ ok: true, action: 'create' });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/**
 * Handles GET requests — used by admin.html to list all orders.
 * Requires ?key=SECRET_KEY to match, or it refuses to return data.
 */
function doGet(e) {
  const key = e.parameter.key;
  if (key !== SECRET_KEY) {
    return jsonResponse({ error: 'Unauthorized — missing or incorrect key.' });
  }

  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values.shift() || COLUMNS;

  const orders = values
    .filter(row => row[1]) // skip blank rows (must have an OrderID)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .reverse(); // newest first

  return jsonResponse({ ok: true, orders: orders });
}

function markOrderPaid(orderId, flwTxRef, flwTransactionId) {
  if (!orderId) return;
  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  const idCol = COLUMNS.indexOf('OrderID');
  const paidCol = COLUMNS.indexOf('PaymentConfirmed');
  const txRefCol = COLUMNS.indexOf('FlwTxRef');
  const txIdCol = COLUMNS.indexOf('FlwTransactionId');

  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === orderId) {
      const rowNum = i + 1; // 1-indexed, +1 for header row already accounted by loop start
      sheet.getRange(rowNum, paidCol + 1).setValue('yes');
      sheet.getRange(rowNum, txRefCol + 1).setValue(flwTxRef || '');
      sheet.getRange(rowNum, txIdCol + 1).setValue(flwTransactionId || '');
      break;
    }
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || createBoundSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Apps Script Web Apps aren't bound to a spreadsheet by default when
// deployed standalone — this creates one on first run if needed.
function createBoundSpreadsheet() {
  const existing = DriveApp.getFilesByName('LeadVault Orders');
  if (existing.hasNext()) {
    return SpreadsheetApp.open(existing.next());
  }
  const ss = SpreadsheetApp.create('LeadVault Orders');
  return ss;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
