/**
 * UNT Error Bank — backend
 *
 * SETUP
 * 1. Go to https://sheets.google.com and create a new blank spreadsheet.
 *    (You can name it "UNT Error Bank Data" — the script will create its
 *    own tabs inside it automatically, you don't need to set up columns.)
 * 2. In the sheet, open Extensions > Apps Script.
 * 3. Delete any starter code in Code.gs and paste this whole file in its place.
 * 4. Click Deploy > New deployment.
 *    - Click the gear icon next to "Select type" and choose "Web app".
 *    - Description: anything, e.g. "UNT Error Bank API".
 *    - Execute as: "Me".
 *    - Who has access: "Anyone".  <-- important, this is what lets students
 *      submit results without a Google account.
 * 5. Click Deploy, authorize the script when Google asks (it's your own
 *    script, this is expected), then copy the "Web app URL" — it ends in /exec.
 * 6. Paste that URL into the app's Settings tab, under "Server URL".
 *
 * That's it — no database setup, no columns to configure. The two sheet tabs
 * ("Tests" and "Results") are created automatically the first time each is used.
 *
 * SECURITY NOTE
 * This deployment (Access: "Anyone") is unauthenticated — anyone who has the
 * URL can read or write to it, the same way anyone with a student test link
 * can currently submit a result. Don't put anything sensitive beyond what's
 * already in the test/results data itself.
 */

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(["id", "json", "updatedAt"]);
  }
  return sheet;
}

function readAll_(sheetName) {
  var sheet = getSheet_(sheetName);
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    try { out.push(JSON.parse(data[i][1])); } catch (e) { /* skip bad row */ }
  }
  return out;
}

function upsert_(sheetName, id, obj) {
  var sheet = getSheet_(sheetName);
  var data = sheet.getDataRange().getValues();
  var json = JSON.stringify(obj);
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 2).setValue(json);
      sheet.getRange(i + 1, 3).setValue(new Date());
      return;
    }
  }
  sheet.appendRow([id, json, new Date()]);
}

function remove_(sheetName, id) {
  var sheet = getSheet_(sheetName);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === id) sheet.deleteRow(i + 1);
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e.parameter.action;
  if (action === "listTests") return jsonOut_({ ok: true, tests: readAll_("Tests") });
  if (action === "listResults") return jsonOut_({ ok: true, results: readAll_("Results") });
  return jsonOut_({ ok: false, error: "unknown action: " + action });
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {
    return jsonOut_({ ok: false, error: "invalid JSON body" });
  }
  var action = body.action;
  if (action === "saveTest") { upsert_("Tests", body.test.id, body.test); return jsonOut_({ ok: true }); }
  if (action === "deleteTest") { remove_("Tests", body.id); return jsonOut_({ ok: true }); }
  if (action === "saveResult") { upsert_("Results", body.result.id, body.result); return jsonOut_({ ok: true }); }
  if (action === "deleteResult") { remove_("Results", body.id); return jsonOut_({ ok: true }); }
  return jsonOut_({ ok: false, error: "unknown action: " + action });
}
