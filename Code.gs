/************************************************************************
 * Durable storage + live slot-blocking for the peer-mentor booking page.
 * Every request is appended to a Google Sheet and read back by the page,
 * so the hidden panel shows all requests and taken slots disappear for
 * everyone.
 *
 * ONE-TIME SETUP (~5 minutes)
 * 1. Go to https://sheets.new  → new Google Sheet. Name it "Mentee bookings".
 * 2. Extensions → Apps Script. Delete the sample, paste this whole file, save.
 * 3. Deploy → New deployment → gear → Web app.
 *      Execute as: Me     Who has access: Anyone     → Deploy.
 *    Authorize (your account → Advanced → "Go to … (unsafe)" → Allow).
 * 4. Copy the Web app URL (ends in /exec).
 * 5. In index.html, paste that URL into the single `endpoint` field in CONFIG.
 *    Commit index.html.
 *
 * If you edit this code later, redeploy: Deploy → Manage deployments →
 * edit (pencil) → Version: New version → Deploy. (Keeps the same URL.)
 ************************************************************************/

var SHEET_NAME = 'Bookings';
var HEADERS = ['Submitted', 'Meeting time', 'Name', 'Email', 'About', 'SlotISO'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEADERS); }
  return sh;
}

function readRows_() {
  var values = sheet_().getDataRange().getValues();
  values.shift(); // header
  return values.map(function (r) {
    return { submittedAt: r[0], when: r[1], name: r[2], email: r[3], about: r[4], slotISO: r[5] };
  });
}

function out_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

// The page uses JSONP (GET) for BOTH listing and saving, so it can read the
// result and detect a slot that was taken a moment earlier.
function doGet(e) {
  var p = (e && e.parameter) || {};
  var callback = p.callback;

  if (p.action === 'save') {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var slot = p.slotISO || '';
      // reject if this exact start time is already booked
      var taken = readRows_().some(function (r) { return r.slotISO === slot; });
      if (taken) return out_({ ok: false, reason: 'taken' }, callback);

      sheet_().appendRow([
        p.submittedAt || new Date().toISOString(),
        p.when || '', p.name || '', p.email || '', p.about || '', slot
      ]);
      return out_({ ok: true }, callback);
    } catch (err) {
      return out_({ ok: false, error: String(err) }, callback);
    } finally {
      lock.releaseLock();
    }
  }

  // default: list newest first
  return out_(readRows_().reverse(), callback);
}

// Fallback path for the fire-and-forget POST (used only if a JSONP save fails).
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var d = JSON.parse(e.postData.contents);
    var taken = readRows_().some(function (r) { return r.slotISO === (d.slotISO || ''); });
    if (!taken) {
      sheet_().appendRow([
        d.submittedAt || new Date().toISOString(),
        d.when || '', d.name || '', d.email || '', d.about || '', d.slotISO || ''
      ]);
    }
    return out_({ ok: !taken });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
