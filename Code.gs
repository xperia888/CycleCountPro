function doGet(e) {
  if (e && e.parameter && e.parameter.health === '1') {
    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      service: 'CycleCount Pro v2 Google Sheet Sync',
      status: 'READY',
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CycleCount Pro v2 — Warehouse Inventory Count & Audit')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var payload = JSON.parse(raw);

    if (payload.action === 'ping') {
      return ContentService.createTextOutput(JSON.stringify({
        ok: true,
        service: 'CycleCount Pro v2',
        message: 'Connected'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var result = appendAuditRecords(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: err && err.message ? err.message : String(err)
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function appendAuditRecords(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var records = Array.isArray(payload.records) ? payload.records : [payload];
    if (!records.length) {
      return { ok: true, saved: 0 };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var headers = [
      'Timestamp', 'Warehouse ID', 'Auditor', 'Pallet LP', 'Inbound ID', 'PO No',
      'Sys Location', 'Actual Location', 'Item ID', 'Item Description',
      'Sys Actual', 'Allocation', 'Remaining', 'Counted Qty', 'Variance',
      'Status', 'Reason', 'Notes'
    ];

    records.forEach(function(r) {
      var whTab = String(r.warehouseId || 'UNASSIGNED').replace(/[^\w-]/g, '_');
      [whTab, 'Master_Audit_Log'].forEach(function(tabName) {
        var sheet = ss.getSheetByName(tabName);
        if (!sheet) {
          sheet = ss.insertSheet(tabName);
          sheet.appendRow(headers);
          sheet.setFrozenRows(1);
          sheet.getRange(1, 1, 1, headers.length)
            .setFontWeight('bold')
            .setBackground('#1e293b')
            .setFontColor('#ffffff');
        }

        var sysActual = Number(r.sysActualQty || 0);
        var counted = Number(r.actualQty || 0);
        var variance = counted - sysActual;

        sheet.appendRow([
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
          r.notes || ''
        ]);
      });
    });

    return { ok: true, saved: records.length };
  } finally {
    lock.releaseLock();
  }
}
