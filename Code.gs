// CycleCount Pro v2 — Google Apps Script backend
//
// Setup:
//   * Preferred: create this script from the target Google Sheet
//     (Extensions -> Apps Script) so getActiveSpreadsheet() works.
//   * Standalone script: set Script Property SPREADSHEET_ID to the Sheet's ID
//     (Project Settings -> Script properties).
//   * The HTML file in this project must be named "index" (matches index.html
//     in the repo). "Index" is also accepted for older copies.

var HTML_FILE_NAMES = ['index', 'Index'];
var MASTER_TAB = 'Master_Audit_Log';
var HEADERS = [
  'Timestamp', 'Warehouse ID', 'Auditor', 'Pallet LP', 'Inbound ID', 'PO No',
  'Sys Location', 'Actual Location', 'Item ID', 'Item Description',
  'Sys Actual', 'Allocation', 'Remaining', 'Counted Qty', 'Variance',
  'Status', 'Reason', 'Notes', 'Record ID'
];
var RECORD_ID_COL = HEADERS.indexOf('Record ID') + 1;

function doGet(e) {
  if (e && e.parameter && e.parameter.health === '1') {
    return jsonOutput_(getSyncHealth());
  }

  return loadTemplate_()
    .evaluate()
    .setTitle('CycleCount Pro v2 — Warehouse Inventory Count & Audit')
    // ALLOWALL is kept so the app can be embedded in Google Sites.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var payload = JSON.parse(raw);

    if (payload.action === 'ping') {
      return jsonOutput_(getSyncHealth());
    }

    return jsonOutput_(appendAuditRecords(payload));
  } catch (err) {
    return jsonOutput_({ ok: false, error: errorMessage_(err) });
  }
}

// Called from the page via google.script.run (when hosted by this script)
// and from doGet/doPost. Confirms the target spreadsheet can be opened.
function getSyncHealth() {
  try {
    var ss = getSpreadsheet_();
    return {
      ok: true,
      service: 'CycleCount Pro v2 Google Sheet Sync',
      status: 'READY',
      spreadsheet: ss.getName(),
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    return { ok: false, status: 'NOT_CONFIGURED', error: errorMessage_(err) };
  }
}

// Called from the page via google.script.run, or from doPost.
// Records carrying a recordId that is already in the Master log are skipped,
// so the client can safely retry a batch whose response it never received.
function appendAuditRecords(payload) {
  payload = payload || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { ok: false, error: 'Sheet is busy (lock timeout). The batch stays queued and will retry.' };
  }
  try {
    var records = Array.isArray(payload.records) ? payload.records : [payload];
    records = records.filter(function(r) { return r && typeof r === 'object'; });
    if (!records.length) {
      return { ok: true, saved: 0, skipped: 0, savedIds: [] };
    }

    var ss = getSpreadsheet_();
    var master = getOrCreateSheet_(ss, MASTER_TAB);
    var existingIds = readExistingIds_(master);

    var rowsByTab = {};
    var savedIds = [];
    var skipped = 0;

    records.forEach(function(r) {
      var recordId = r.recordId ? String(r.recordId) : '';
      if (recordId && existingIds[recordId]) {
        skipped++;
        savedIds.push(recordId); // already stored earlier: report as confirmed
        return;
      }
      if (recordId) existingIds[recordId] = true;

      var row = buildRow_(r, recordId);
      var whTab = String(r.warehouseId || 'UNASSIGNED').replace(/[^\w-]/g, '_');
      [whTab, MASTER_TAB].forEach(function(tabName) {
        (rowsByTab[tabName] = rowsByTab[tabName] || []).push(row);
      });
      if (recordId) savedIds.push(recordId);
    });

    Object.keys(rowsByTab).forEach(function(tabName) {
      var sheet = tabName === MASTER_TAB ? master : getOrCreateSheet_(ss, tabName);
      var rows = rowsByTab[tabName];
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    });
    SpreadsheetApp.flush();

    return {
      ok: true,
      saved: records.length - skipped,
      skipped: skipped,
      savedIds: savedIds
    };
  } catch (err) {
    return { ok: false, error: errorMessage_(err) };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function loadTemplate_() {
  var lastErr;
  for (var i = 0; i < HTML_FILE_NAMES.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(HTML_FILE_NAMES[i]);
    } catch (err) {
      lastErr = err;
    }
  }
  throw new Error('No HTML file named "index" found in this Apps Script project. ' + errorMessage_(lastErr));
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('No spreadsheet linked. Create the script from the Sheet (Extensions -> Apps Script) ' +
    'or set the SPREADSHEET_ID script property.');
}

function getOrCreateSheet_(ss, tabName) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    styleHeader_(sheet);
    return sheet;
  }
  // Upgrade tabs created by the earlier version (no Record ID column).
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }
  var headerCell = sheet.getRange(1, RECORD_ID_COL);
  if (headerCell.getValue() !== 'Record ID') {
    headerCell.setValue('Record ID');
    styleHeader_(sheet);
  }
  return sheet;
}

function styleHeader_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff');
}

function readExistingIds_(sheet) {
  var ids = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ids;
  var values = sheet.getRange(2, RECORD_ID_COL, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) ids[String(values[i][0])] = true;
  }
  return ids;
}

function buildRow_(r, recordId) {
  var sysActual = Number(r.sysActualQty || 0);
  var counted = Number(r.actualQty || 0);
  var variance = counted - sysActual;
  return [
    r.countedAt || new Date().toISOString(),
    r.warehouseId || '',
    r.auditor || '',
    r.lpn || '',
    r.inboundNo || '',
    r.poNo || '',
    r.sysLocation || '',
    r.actualLocation || '',
    r.itemId || '',
    r.itemName || '',
    sysActual,
    Number(r.sysAllocQty || 0),
    Number(r.sysAvailQty || 0),
    counted,
    variance,
    r.status || (variance === 0 ? 'Matched' : 'Variance'),
    r.reasonCode || '',
    r.notes || '',
    recordId || ''
  ];
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorMessage_(err) {
  return err && err.message ? err.message : String(err);
}
