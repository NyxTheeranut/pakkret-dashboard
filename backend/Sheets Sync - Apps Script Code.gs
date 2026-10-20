/**
 * Pak Kret Territory Explorer -- Google Sheets backend (viewer read access).
 *
 * What this is: the API the hosted Territory Explorer page calls so a viewer
 * only has to sign in with Google -- no file picker, no per-viewer Drive
 * sharing to manage. It runs inside the "PakKret Data" Sheet, as that Sheet's
 * owner -- there is no separate server to host or pay for, and the Sheet
 * itself is the database. Modeled on the Route Planner's Apps Script backend
 * (see the Route Planner repo's "Sheets Sync - Apps Script Code.gs").
 *
 * ── Why this exists ──────────────────────────────────────────────────────
 * The dashboard used to have viewers sign in with an OAuth `drive.file`
 * token and pick the spreadsheet from a Picker before it could be read via
 * the Sheets REST API. That's not just a UX step: under `drive.file` scope,
 * ordinary Drive "Share" access does NOT by itself let the app read a file --
 * only running that exact file through the Picker records the per-file
 * grant. A viewer who never used the Picker got a 404 even though the sheet
 * was shared with them.
 *
 * This script sidesteps that entirely: it runs AS the sheet's owner (a bound
 * script, deployed "Execute as: Me"), so it already has permanent access.
 * The viewer only proves who they are with a lightweight Google Identity
 * Services sign-in (an ID token, not an API access token) -- no Drive/Sheets
 * OAuth scope is ever requested from them. Access control is an email
 * allowlist in this sheet instead of Drive sharing.
 *
 * ── Data layout ────────────────────────────────────────────────────────────
 * "meta", "rtr-YYYY-MM-*", "sev-YYYY-MM-*" -- written by
 *   update_territory_sheet.py (a plain local script, no browser) via the
 *   publishTerritory action below. This script only reads them back, by tab
 *   name, exactly as before.
 * "Viewers" -- who's allowed to read data. One column: email. Created
 *   automatically (with a sample row) the first time anyone signs in, if it
 *   doesn't exist yet -- replace the sample row with real teammate emails.
 *
 * ── Auth (viewers) ─────────────────────────────────────────────────────────
 * The page signs the user in with Google Identity Services and sends the
 * resulting ID token on every request. This script verifies that token
 * against Google directly (no session/cookie trust needed) and checks the
 * token's audience against OAUTH_CLIENT_ID, so it only accepts tokens issued
 * for THIS app. The verified email is then checked against the Viewers tab.
 * getValuesBatch runs this same check, just once for a whole list of tabs
 * instead of once per tab -- see its own comment for why (Apps Script's
 * simultaneous-execution quota, not extra security).
 *
 * ── Auth (publishing) ──────────────────────────────────────────────────────
 * publishTerritory and getSyncData aren't a person signing in -- they're
 * update_territory_sheet.py on the admin's own machine -- so neither can go
 * through the Viewers tab. Gated by a shared secret instead
 * (TERRITORY_SYNC_SECRET), same pattern as Route Planner's
 * syncStores/requireSyncSecret_. Neither endpoint does any business logic of
 * its own: getSyncData just hands back the raw tabs/meta as they currently
 * are, and publishTerritory writes exactly the tabs it's given and deletes
 * exactly the tab names it's told to -- every actual decision (targets, CM
 * rollups, which months survive retention, which months get merged in from
 * the Sheet's existing state) is made in the Python script, not here.
 *
 * ── SETUP (one-time) ─────────────────────────────────────────────────────
 *  1. Open the "PakKret Data" Sheet. Extensions -> Apps Script. Delete any
 *     starter code, paste this whole file in.
 *  2. Project Settings (gear icon, left sidebar) -> Script Properties -> Add:
 *       OAUTH_CLIENT_ID = 901392336998-4hkkmokj5mtph2u4c2dtem7mplv3o7sv.apps.googleusercontent.com
 *       (the same public Client ID already embedded in the dashboard's HTML
 *       -- it's already an OAuth "Web application" client with this site's
 *       origins authorized, and Identity Services sign-in can share it.)
 *       TERRITORY_SYNC_SECRET = <any random string -- generate one with, e.g.,
 *       `openssl rand -hex 24`. Also put this exact value into
 *       Config/territory_sync_secret.txt, which update_territory_sheet.py
 *       reads. Without this, "Anyone" access below means anyone, not just
 *       your own script, could overwrite the published data.>
 *  3. Deploy -> New deployment -> Type: Web app.
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     ("Anyone" is fine here -- real access control happens via the ID token
 *     + Viewers tab check (reads) and the shared secret (publish), not via
 *     this deployment setting.)
 *  4. Deploy, authorize when prompted, copy the Web app URL into the page's
 *     VIEWER_SCRIPT_URL config AND into update_territory_sheet.py's SYNC_URL.
 *  5. Sign in once from the page -- this creates the Viewers tab. Edit the
 *     sample row to a real email, add the rest of the team below it.
 *  6. Optional cleanup: viewers no longer need the file Drive-shared with
 *     them individually -- the Viewers tab is now the real gate.
 */

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.action === "getValues") {
      return jsonResponse_(getValuesForViewer_(body.idToken, body.tab));
    }

    if (body.action === "getValuesBatch") {
      // Same viewer auth as getValues, but returns MULTIPLE tabs in one response -- see
      // getValuesBatchForViewer_'s own comment for why this exists.
      return jsonResponse_(getValuesBatchForViewer_(body.idToken, body.tabs));
    }

    if (body.action === "publishTerritory") {
      requireSyncSecret_(body.secret);
      return jsonResponse_(publishTerritory_(body));
    }

    if (body.action === "getSyncData") {
      // Read-only counterpart of publishTerritory, for update_territory_sheet.py to fetch
      // the Sheet's current state before merging in fresh local months -- gated by the same
      // secret (not a person signing in, so no ID token). See getSyncData_'s own comment for
      // why this round-trip exists.
      requireSyncSecret_(body.secret);
      return jsonResponse_(getSyncData_());
    }

    return jsonResponse_({ ok: false, error: "unknown action" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// update_territory_sheet.py isn't a person signing in -- it's a script on the
// admin's own machine -- so it can't be checked against the Viewers tab. A
// shared secret (set once as a Script Property, and passed by the python
// script) is the minimum needed so this deployment's public URL alone isn't
// enough to overwrite the published data.
function requireSyncSecret_(secret) {
  var expected = PropertiesService.getScriptProperties().getProperty(
    "TERRITORY_SYNC_SECRET",
  );
  if (!expected)
    throw new Error(
      "TERRITORY_SYNC_SECRET script property is not set -- see setup notes at the top of this file",
    );
  if (secret !== expected) throw new Error("forbidden: bad sync secret");
}

// ---------- publishTerritory: dumb executor, no business logic ----------
// payload = {
//   tables: { "<tabName>": { headers: [...], rows: [[...], ...] }, ... },
//   meta: { rtr_months: [...], sev_months: [...], updated_at: <ms> },
//   deleteTabs: ["<tabName>", ...],
// }
// Writes each table as a plain grid (identical shape to what the browser's
// own writeTableTab already wrote -- the read side, scriptReadTable, is
// agnostic to how a tab was populated). Writes meta as a single JSON string
// in cell A1 of the "meta" tab -- scriptReadTab already does
// values.map(r=>r[0]).join("") then JSON.parse, and a join-of-one is exactly
// that string, so no viewer-side change is needed for this. Deletes every
// tab name in deleteTabs -- retention/pruning is the Python script's
// decision, this just carries it out.
function publishTerritory_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tables = payload.tables || {};
  Object.keys(tables).forEach(function (tabName) {
    var t = tables[tabName];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) sheet = ss.insertSheet(tabName);
    sheet.clearContents();
    var grid = [t.headers].concat(t.rows);
    if (grid.length && grid[0].length) {
      var range = sheet.getRange(1, 1, grid.length, grid[0].length);
      // Set BEFORE writing, not after: Range.setValues() does the same locale-aware
      // auto-detection as USER_ENTERED input (unlike the old direct-REST-API sync, which
      // explicitly used RAW to avoid this) -- a label like "May 2026" or "Aug.26" gets
      // silently parsed into a real date, which then serializes to a full ISO timestamp
      // ("2026-05-01T00:00:00.000Z") the instant a viewer reads it back as JSON. This is
      // the exact same bug class as Route Planner's "7-11 got parsed as a date" fix --
      // forcing the whole range to plain text before the values land avoids it entirely,
      // regardless of which specific column would otherwise look date-like.
      range.setNumberFormat("@");
      range.setValues(grid);
    }
  });

  if (payload.meta) {
    var metaSheet = ss.getSheetByName("meta");
    if (!metaSheet) metaSheet = ss.insertSheet("meta");
    metaSheet.clearContents();
    var metaRange = metaSheet.getRange(1, 1);
    metaRange.setNumberFormat("@");
    metaRange.setValue(JSON.stringify(payload.meta));
  }

  (payload.deleteTabs || []).forEach(function (tabName) {
    var sheet = ss.getSheetByName(tabName);
    if (sheet) ss.deleteSheet(sheet);
  });

  return {
    ok: true,
    tablesWritten: Object.keys(tables).length,
    tabsDeleted: (payload.deleteTabs || []).length,
  };
}

// ---------- getSyncData: update_territory_sheet.py's read-before-merge ----------
// Returns every existing rtr-/sev- tab as a raw grid, plus meta, in the exact same
// {tables: {tabName: {headers, rows}}, meta: {...}} shape publishTerritory_ writes --
// so update_territory_sheet.py can merge fresh local months into this directly with no
// reconstruction step (unlike Route Planner/TOL Tracker, this project's storage is
// already flat per-month tabs, not one nested blob, so there's no nested shape to rebuild).
// Exists so a local run missing an older source file can't make that month silently
// vanish from the dashboard -- see build_payload's own docstring in the Python script.
function getSyncData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tables = {};
  var meta = null;
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (name === "meta") {
      var metaValues = sheet.getDataRange().getValues();
      var joined = metaValues
        .map(function (r) {
          return r[0];
        })
        .join("");
      try {
        meta = JSON.parse(joined);
      } catch (e) {
        meta = null;
      }
      return;
    }
    if (name.indexOf("rtr-") !== 0 && name.indexOf("sev-") !== 0) return;
    var values = sheet.getDataRange().getValues();
    if (!values.length || !values[0].length) return;
    tables[name] = { headers: values[0], rows: values.slice(1) };
  });
  return { ok: true, tables: tables, meta: meta };
}

// ---------- auth ----------

// Verifies the ID token directly against Google (not just trusting the client) and
// checks it was issued for THIS app specifically, via the audience claim.
function verifyIdToken_(idToken) {
  if (!idToken) return null;
  var resp = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" +
      encodeURIComponent(idToken),
    { muteHttpExceptions: true },
  );
  if (resp.getResponseCode() !== 200) return null;
  var data = JSON.parse(resp.getContentText());
  var expectedClientId =
    PropertiesService.getScriptProperties().getProperty("OAUTH_CLIENT_ID");
  if (!expectedClientId)
    throw new Error(
      "OAUTH_CLIENT_ID script property is not set -- see setup notes at the top of this file",
    );
  if (data.aud !== expectedClientId) return null;
  if (!data.email || data.email_verified !== "true") return null;
  return data.email;
}

// Looks up whether a verified email is on the allowlist. Creates the Viewers tab
// (with a sample row) on first use if it doesn't exist yet, so there's something to edit.
function lookupViewer_(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Viewers");
  if (!sheet) {
    sheet = ss.insertSheet("Viewers");
    sheet.appendRow(["email"]);
    sheet.appendRow(["example@gmail.com"]);
    sheet.setFrozenRows(1);
    return false; // just created -- nothing real to match against yet
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase())
      return true;
  }
  return false;
}

// ---------- getValues: the whole point of this script ----------
// Returns one tab's raw grid, in the same {values: [[...]]} shape the Sheets
// REST API's values.get already returns -- so the dashboard's existing
// per-row parsing logic (readRtrMonthGroup, readSevMonthGroup, ...) doesn't
// need to change at all, only the transport that feeds it.
function getValuesForViewer_(idToken, tab) {
  var email = verifyIdToken_(idToken);
  if (!email) return { ok: false, error: "not_signed_in" };
  if (!lookupViewer_(email)) {
    return {
      ok: false,
      error: "no_access",
      message:
        "บัญชีนี้ (" +
        email +
        ") ยังไม่ได้รับสิทธิ์ดูข้อมูล -- แจ้งผู้ดูแลให้เพิ่มอีเมลนี้ในแท็บ Viewers",
    };
  }
  if (!tab || !isReadableTab_(tab)) return { ok: false, error: "no_such_tab" };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
  if (!sheet) return { ok: false, error: "no_such_tab" };
  return { ok: true, values: sheet.getDataRange().getValues() };
}

// ---------- getValuesBatch: same viewer auth as getValues, but for MULTIPLE tabs ----------
// loadFromCloud used to fire one getValues call PER TAB -- ~54 for a full 6-month load --
// comfortably over Apps Script's own simultaneous-execution quota (documented around 30
// for a personal Google account) the moment more than a couple of viewers load around the
// same time; past that ceiling Apps Script starts returning degraded/error responses that
// often carry no CORS header at all, which the browser reports as a misleading "blocked by
// CORS policy" instead of the real cause. Doing the auth check ONCE here and returning
// every requested tab's raw grid together cuts a full page load down to about 2 requests
// total (one for meta, one for everything else), regardless of how many months that
// covers -- so the simultaneous-execution ceiling is no longer in reach even with several
// viewers loading at once. Per-tab results use the exact same {ok, values} / {ok, error}
// shape getValues already returns, keyed by tab name, so scriptGetValuesBatch on the
// dashboard side can reuse getValues' own error handling per tab.
function getValuesBatchForViewer_(idToken, tabs) {
  var email = verifyIdToken_(idToken);
  if (!email) return { ok: false, error: "not_signed_in" };
  if (!lookupViewer_(email)) {
    return {
      ok: false,
      error: "no_access",
      message:
        "บัญชีนี้ (" +
        email +
        ") ยังไม่ได้รับสิทธิ์ดูข้อมูล -- แจ้งผู้ดูแลให้เพิ่มอีเมลนี้ในแท็บ Viewers",
    };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  (tabs || []).forEach(function (tab) {
    if (!tab || !isReadableTab_(tab)) {
      result[tab] = { ok: false, error: "no_such_tab" };
      return;
    }
    var sheet = ss.getSheetByName(tab);
    if (!sheet) {
      result[tab] = { ok: false, error: "no_such_tab" };
      return;
    }
    result[tab] = { ok: true, values: sheet.getDataRange().getValues() };
  });
  return { ok: true, tabs: result };
}

// Being on the Viewers allowlist only proves WHO someone is, not that every tab in the
// spreadsheet is meant for them -- without this, an approved viewer could ask for the
// "Viewers" tab itself and read every other approved teammate's email. Only the tabs the
// dashboard actually reads are servable; everything else (including "Viewers") answers
// the same "no_such_tab" as a genuinely missing tab, so this doesn't even reveal that a
// tab exists but is off-limits.
function isReadableTab_(tab) {
  return (
    tab === "meta" || tab.indexOf("rtr-") === 0 || tab.indexOf("sev-") === 0
  );
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
