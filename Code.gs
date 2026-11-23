// CycleCount Pro v2 — Google Apps Script backend
//
// Setup (Project Settings -> Script properties):
//   * SPREADSHEET_ID  only if this script was not created from the Sheet
//                     (Extensions -> Apps Script); otherwise not needed.
//   * MONITOR_CODE    access code for the manager monitor page (?view=monitor).
//                     Use a long code; changing it signs out every monitor screen.
// Times written by versions before Nov 2026 were UTC text: run fixSheetTimes once.
// Monthly archive (keeps the last 90 days in this Sheet): select setupMonthlyArchive
// in the editor and press Run once, then approve the permissions it asks for.
// HTML files in the project: "index" (counter app) and "monitor" (monitor page).

var HTML_FILE_NAMES = ['index', 'Index'];
var MONITOR_FILE_NAMES = ['monitor', 'Monitor'];
var MASTER_TAB = 'Master_Audit_Log';
var TASK_TAB = 'Count_Tasks';
var EXPECTED_TAB = 'Task_Expected';

var HEADERS = [
  'Timestamp', 'Warehouse ID', 'Auditor', 'Pallet LP', 'Inbound ID', 'PO No',
  'Sys Location', 'Actual Location', 'Item ID', 'Item Description',
  'Sys Actual', 'Allocation', 'Remaining', 'Counted Qty', 'Variance',
  'Status', 'Reason', 'Notes', 'Record ID', 'Task ID', 'Customer No', 'Customer Name'
];
var COL = {};
HEADERS.forEach(function(h, i) { COL[h] = i; });
var RECORD_ID_COL = COL['Record ID'] + 1;

var TASK_HEADERS = [
  'Task ID', 'Warehouse ID', 'Customer No', 'Customer Name', 'Locations', 'Assigned To',
  'Status', 'Source', 'Created At', 'Created By', 'Started At', 'Completed At',
  'Expected Pallets', 'Notes'
];
var TCOL = {};
TASK_HEADERS.forEach(function(h, i) { TCOL[h] = i; });

var EXPECTED_HEADERS = [
  'Task ID', 'Warehouse ID', 'Pallet LP', 'Item ID', 'Item Description', 'Sys Location',
  'Sys Actual', 'Customer No', 'Uploaded At'
];

var OPEN_STATUSES = ['Open', 'In Progress'];

// Times are stored as real date-times (shown in the Sheet's time zone, File > Settings),
// not as UTC text, so they read as local time and sort and filter as dates.
var DATE_TIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';
var SHEET_TIME_ZONE = 'America/Los_Angeles';
var TASK_TIME_COLS = ['Created At', 'Started At', 'Completed At'];

// Bumped when the JSON API changes. Clients check it (GET ?health=1) before sending
// actions, so an outdated deployment is reported instead of misreading requests.
var API_VERSION = 3;

// ---------------------------------------------------------------------------
// Web app entry points
// ---------------------------------------------------------------------------

function doGet(e) {
  var params = (e && e.parameter) || {};
  if (params.health === '1') {
    return jsonOutput_(getSyncHealth());
  }

  var isMonitor = params.view === 'monitor';
  return loadTemplate_(isMonitor ? MONITOR_FILE_NAMES : HTML_FILE_NAMES)
    .evaluate()
    .setTitle(isMonitor ? 'CycleCount Pro — Monitoring Center' : 'CycleCount Pro v2 — Warehouse Inventory Count & Audit')
    // ALLOWALL is kept so the app can be embedded in Google Sites.
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// JSON API for clients that are not hosted by this script (fetch POST, text/plain body).
var POST_ACTIONS = {
  appendAuditRecords: function(p) { return appendAuditRecords(p); },
  listTasks: function(p) { return listTasks(p); },
  createTask: function(p) { return createTask(p); },
  startTask: function(p) { return startTask(p); },
  completeTask: function(p) { return completeTask(p); },
  getMonitorData: function(p) { return getMonitorData(p); },
  monitorCreateTask: function(p) { return monitorCreateTask(p); },
  monitorUpdateTask: function(p) { return monitorUpdateTask(p); },
  monitorArchive: function(p) { return monitorArchive(p); }
};

function doPost(e) {
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var payload = JSON.parse(raw);

    if (payload.action === 'ping') {
      return jsonOutput_(getSyncHealth());
    }
    if (payload.action) {
      if (POST_ACTIONS[payload.action]) return jsonOutput_(POST_ACTIONS[payload.action](payload.payload || {}));
      return jsonOutput_({ ok: false, error: 'Unknown action: ' + payload.action });
    }
    // No action (older clients): payload is an audit batch.
    return jsonOutput_(appendAuditRecords(payload));
  } catch (err) {
    return jsonOutput_({ ok: false, error: errorMessage_(err) });
  }
}

// Confirms the target spreadsheet can be opened.
function getSyncHealth() {
  try {
    var ss = getSpreadsheet_();
    return {
      ok: true,
      service: 'CycleCount Pro v2 Google Sheet Sync',
      status: 'READY',
      apiVersion: API_VERSION,
      spreadsheet: ss.getName(),
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    return { ok: false, status: 'NOT_CONFIGURED', apiVersion: API_VERSION, error: errorMessage_(err) };
  }
}

// ---------------------------------------------------------------------------
// Count records
// ---------------------------------------------------------------------------

// Records carrying a recordId that is already in the Master log are skipped,
// so the client can safely retry a batch whose response it never received.
function appendAuditRecords(payload) {
  payload = payload || {};
  return withLock_(function() {
    var records = Array.isArray(payload.records) ? payload.records : [payload];
    records = records.filter(function(r) { return r && typeof r === 'object' && (r.lpn || r.recordId); });
    if (!records.length) {
      return { ok: true, saved: 0, skipped: 0, savedIds: [] };
    }

    var ss = getSpreadsheet_();
    var master = getOrCreateSheet_(ss, MASTER_TAB, HEADERS);
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
      var whTab = safeTabName_(r.warehouseId || 'UNASSIGNED');
      [whTab, MASTER_TAB].forEach(function(tabName) {
        (rowsByTab[tabName] = rowsByTab[tabName] || []).push(row);
      });
      if (recordId) savedIds.push(recordId);
    });

    Object.keys(rowsByTab).forEach(function(tabName) {
      var sheet = tabName === MASTER_TAB ? master : getOrCreateSheet_(ss, tabName, HEADERS);
      var rows = rowsByTab[tabName];
      var start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
      sheet.getRange(start, COL['Timestamp'] + 1, rows.length, 1).setNumberFormat(DATE_TIME_FORMAT);
    });
    SpreadsheetApp.flush();
    clearMonitorCache_();

    return { ok: true, saved: records.length - skipped, skipped: skipped, savedIds: savedIds };
  });
}

// ---------------------------------------------------------------------------
// Count tasks (used by the counter app)
// ---------------------------------------------------------------------------

// Open / In Progress tasks in a warehouse that are unassigned or assigned to this counter.
function listTasks(params) {
  params = params || {};
  try {
    var ss = getSpreadsheet_();
    var wh = normWh_(params.warehouseId);
    // counter may be one name or several (WMS username and display name).
    var me = (Array.isArray(params.counter) ? params.counter : [params.counter])
      .map(normName_).filter(Boolean);
    var tasks = readTasks_(ss).filter(function(t) {
      if (wh && normWh_(t.warehouseId) !== wh) return false;
      if (OPEN_STATUSES.indexOf(t.status) === -1) return false;
      var assigned = normName_(t.assignedTo);
      return !assigned || me.indexOf(assigned) !== -1;
    });
    return { ok: true, tasks: tasks };
  } catch (err) {
    return { ok: false, error: errorMessage_(err) };
  }
}

// Counter creates their own task (assigned to themselves, In Progress).
function createTask(params) {
  params = params || {};
  return withLock_(function() {
    if (!params.warehouseId) throw new Error('Warehouse is required.');
    if (!params.customerNo) throw new Error('Customer number is required.');
    var counter = String(params.counter || '').trim();
    var task = addTask_(getSpreadsheet_(), {
      warehouseId: params.warehouseId,
      customerNo: params.customerNo,
      customerName: params.customerName,
      locations: params.locations,
      assignedTo: counter,
      status: 'In Progress',
      source: 'Counter',
      createdBy: counter || 'Counter',
      startedAt: new Date(),
      notes: params.notes
    });
    clearMonitorCache_();
    return { ok: true, task: task };
  });
}

// Called after the counter loads the task's inventory: stores the expected pallet
// list (replacing any earlier upload for this task) and marks the task In Progress.
function startTask(params) {
  params = params || {};
  return withLock_(function() {
    var ss = getSpreadsheet_();
    var found = findTask_(ss, params.taskId);
    if (!found) throw new Error('Task ' + params.taskId + ' not found.');
    var t = found.task;
    var items = Array.isArray(params.items) ? params.items : [];
    var now = new Date();

    var expSheet = getOrCreateSheet_(ss, EXPECTED_TAB, EXPECTED_HEADERS);
    var lastRow = expSheet.getLastRow();
    var keep = [];
    if (lastRow >= 2) {
      keep = expSheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).getValues()
        .filter(function(r) { return String(r[0]) !== t.taskId; });
      expSheet.getRange(2, 1, lastRow - 1, EXPECTED_HEADERS.length).clearContent();
    }
    var fresh = items.map(function(i) {
      return [t.taskId, t.warehouseId, i.lpn || '', i.itemId || '', i.itemName || '', i.sysLocation || '',
        Number(i.sysActualQty || 0), i.customerNo || t.customerNo, now];
    });
    var all = keep.concat(fresh);
    if (all.length) {
      expSheet.getRange(2, 1, all.length, EXPECTED_HEADERS.length).setValues(all.map(function(r) {
        r[8] = sheetTime_(r[8]); return r;
      }));
      expSheet.getRange(2, 9, all.length, 1).setNumberFormat(DATE_TIME_FORMAT);
    }

    var updates = { 'Status': 'In Progress', 'Expected Pallets': fresh.length };
    if (!t.startedAt) updates['Started At'] = now;
    if (!t.assignedTo && params.counter) updates['Assigned To'] = String(params.counter).trim();
    writeTaskFields_(found.sheet, found.row, updates);
    SpreadsheetApp.flush();
    clearMonitorCache_();
    return { ok: true, task: findTask_(ss, t.taskId).task };
  });
}

function completeTask(params) {
  params = params || {};
  return withLock_(function() {
    var ss = getSpreadsheet_();
    var found = findTask_(ss, params.taskId);
    if (!found) throw new Error('Task ' + params.taskId + ' not found.');
    writeTaskFields_(found.sheet, found.row, { 'Status': 'Completed', 'Completed At': new Date() });
    clearMonitorCache_();
    return { ok: true, task: findTask_(ss, params.taskId).task };
  });
}

// ---------------------------------------------------------------------------
// Monitor (manager) API — every call needs the MONITOR_CODE access code
// ---------------------------------------------------------------------------

function getMonitorData(params) {
  params = params || {};
  try {
    checkMonitorCode_(params.code);
    var cache = CacheService.getScriptCache();
    var cacheKey = 'monitor_v1';
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    var result = buildMonitorData_(getSpreadsheet_());
    try { cache.put(cacheKey, JSON.stringify(result), 15); } catch (e) { /* too large to cache */ }
    return result;
  } catch (err) {
    return { ok: false, error: errorMessage_(err), authError: !!(err && err.authError) };
  }
}

// Supervisor assigns a task from the monitor page.
function monitorCreateTask(params) {
  params = params || {};
  try {
    checkMonitorCode_(params.code);
  } catch (err) {
    return { ok: false, error: errorMessage_(err), authError: true };
  }
  return withLock_(function() {
    if (!params.warehouseId) throw new Error('Warehouse is required.');
    if (!params.customerNo) throw new Error('Customer number is required.');
    var task = addTask_(getSpreadsheet_(), {
      warehouseId: params.warehouseId,
      customerNo: params.customerNo,
      customerName: params.customerName,
      locations: params.locations,
      assignedTo: params.assignedTo,
      status: 'Open',
      source: 'Supervisor',
      createdBy: 'Monitor',
      notes: params.notes
    });
    clearMonitorCache_();
    return { ok: true, task: task };
  });
}

// Supervisor reassigns / reopens / cancels / completes a task.
function monitorUpdateTask(params) {
  params = params || {};
  try {
    checkMonitorCode_(params.code);
  } catch (err) {
    return { ok: false, error: errorMessage_(err), authError: true };
  }
  return withLock_(function() {
    var ss = getSpreadsheet_();
    var found = findTask_(ss, params.taskId);
    if (!found) throw new Error('Task ' + params.taskId + ' not found.');
    var updates = {};
    if (params.status) {
      var allowed = ['Open', 'In Progress', 'Completed', 'Cancelled'];
      if (allowed.indexOf(params.status) === -1) throw new Error('Unknown status ' + params.status);
      updates['Status'] = params.status;
      if (params.status === 'Completed') updates['Completed At'] = new Date();
      if (params.status === 'Open' || params.status === 'In Progress') updates['Completed At'] = '';
    }
    if (params.assignedTo !== undefined) updates['Assigned To'] = String(params.assignedTo || '').trim();
    writeTaskFields_(found.sheet, found.row, updates);
    clearMonitorCache_();
    return { ok: true, task: findTask_(ss, params.taskId).task };
  });
}

function buildMonitorData_(ss) {
  var now = Date.now();
  var tasks = readTasks_(ss);
  var taskById = {};
  tasks.forEach(function(t) {
    t.expectedLpns = {};
    t.expected = Number(t.expectedPallets || 0);
    t.counted = 0; t.countedExpected = 0; t.extra = 0;
    t.variances = 0; t.netUnits = 0; t.lastActivity = ''; t.counters = {};
    taskById[t.taskId] = t;
  });

  // Expected pallets per task
  var exp = ss.getSheetByName(EXPECTED_TAB);
  if (exp && exp.getLastRow() >= 2) {
    exp.getRange(2, 1, exp.getLastRow() - 1, 3).getValues().forEach(function(r) {
      var t = taskById[String(r[0])];
      if (t) t.expectedLpns[String(r[2]).toUpperCase()] = true;
    });
    tasks.forEach(function(t) {
      var n = Object.keys(t.expectedLpns).length;
      if (n) t.expected = n;
    });
  }

  // Latest count per (task, warehouse, pallet), plus activity
  var master = ss.getSheetByName(MASTER_TAB);
  var rows = [];
  if (master && master.getLastRow() >= 2) {
    var width = Math.min(HEADERS.length, master.getLastColumn());
    rows = master.getRange(2, 1, master.getLastRow() - 1, width).getValues();
  }
  var latest = {};
  var perPallet = {};
  var counters = {};
  var recent = [];
  var startOfDay = new Date(); startOfDay.setHours(0, 0, 0, 0);

  rows.forEach(function(r) {
    var rec = {
      ts: toIso_(r[COL['Timestamp']]),
      warehouseId: String(r[COL['Warehouse ID']] || ''),
      auditor: String(r[COL['Auditor']] || ''),
      lpn: String(r[COL['Pallet LP']] || '').toUpperCase(),
      sysLocation: String(r[COL['Sys Location']] || ''),
      actualLocation: String(r[COL['Actual Location']] || ''),
      itemId: String(r[COL['Item ID']] || ''),
      itemName: String(r[COL['Item Description']] || ''),
      sysActual: Number(r[COL['Sys Actual']] || 0),
      counted: Number(r[COL['Counted Qty']] || 0),
      variance: Number(r[COL['Variance']] || 0),
      reason: String(r[COL['Reason']] || ''),
      notes: String(r[COL['Notes']] || ''),
      taskId: String(r[COL['Task ID']] || ''),
      customerNo: String(r[COL['Customer No']] || ''),
      customerName: String(r[COL['Customer Name']] || '')
    };
    if (!rec.lpn) return;
    latest[rec.taskId + '|' + rec.warehouseId + '|' + rec.lpn] = rec;

    var pk = rec.warehouseId + '|' + rec.lpn;
    var p = perPallet[pk] = perPallet[pk] || { warehouseId: rec.warehouseId, lpn: rec.lpn, itemId: rec.itemId, tasks: {}, auditors: {}, counts: [] };
    p.tasks[rec.taskId || '(no task)'] = true;
    p.auditors[rec.auditor || '(unknown)'] = true;
    p.counts.push({ auditor: rec.auditor, taskId: rec.taskId, counted: rec.counted, ts: rec.ts });

    var c = counters[rec.auditor || '(unknown)'] = counters[rec.auditor || '(unknown)'] ||
      { name: rec.auditor || '(unknown)', lastActivity: '', countsToday: 0, warehouseId: '', taskId: '', customerNo: '' };
    if (!c.lastActivity || rec.ts > c.lastActivity) {
      c.lastActivity = rec.ts; c.warehouseId = rec.warehouseId; c.taskId = rec.taskId; c.customerNo = rec.customerNo;
    }
    if (rec.ts && new Date(rec.ts).getTime() >= startOfDay.getTime()) c.countsToday++;
    recent.push(rec);
  });

  var variances = [];
  Object.keys(latest).forEach(function(k) {
    var rec = latest[k];
    var t = taskById[rec.taskId];
    var locMismatch = rec.actualLocation && rec.sysLocation && rec.actualLocation !== rec.sysLocation;
    if (t) {
      t.counted++;
      if (t.expectedLpns[rec.lpn]) t.countedExpected++; else t.extra++;
      if (rec.variance !== 0 || locMismatch) t.variances++;
      t.netUnits += rec.variance;
      if (!t.lastActivity || rec.ts > t.lastActivity) t.lastActivity = rec.ts;
      t.counters[rec.auditor] = true;
    }
    if (rec.variance !== 0 || locMismatch) variances.push(rec);
  });
  variances.sort(function(a, b) { return a.ts < b.ts ? 1 : -1; });

  var duplicates = [];
  Object.keys(perPallet).forEach(function(k) {
    var p = perPallet[k];
    var taskIds = Object.keys(p.tasks);
    var auditors = Object.keys(p.auditors);
    if (taskIds.length > 1 || auditors.length > 1) {
      duplicates.push({ warehouseId: p.warehouseId, lpn: p.lpn, itemId: p.itemId, tasks: taskIds, auditors: auditors, counts: p.counts.slice(-6) });
    }
  });

  recent.sort(function(a, b) { return a.ts < b.ts ? 1 : -1; });

  var outTasks = tasks.map(function(t) {
    return {
      taskId: t.taskId, warehouseId: t.warehouseId, customerNo: t.customerNo, customerName: t.customerName,
      locations: t.locations, assignedTo: t.assignedTo, status: t.status, source: t.source,
      createdAt: t.createdAt, startedAt: t.startedAt, completedAt: t.completedAt,
      expected: t.expected, counted: t.counted, countedExpected: t.countedExpected, extra: t.extra,
      variances: t.variances, netUnits: t.netUnits, lastActivity: t.lastActivity,
      counters: Object.keys(t.counters).filter(Boolean)
    };
  });

  return {
    ok: true,
    generatedAt: new Date(now).toISOString(),
    tasks: outTasks,
    counters: Object.keys(counters).map(function(k) { return counters[k]; })
      .sort(function(a, b) { return a.lastActivity < b.lastActivity ? 1 : -1; }),
    variances: variances.slice(0, 300),
    duplicates: duplicates.slice(0, 200),
    recent: recent.slice(0, 40),
    totals: { records: rows.length, varianceCount: variances.length, duplicateCount: duplicates.length },
    archive: lastArchiveInfo_()
  };
}

function checkMonitorCode_(code) {
  var props = PropertiesService.getScriptProperties();
  var expected = props.getProperty('MONITOR_CODE');
  var fail = function(msg) { var e = new Error(msg); e.authError = true; throw e; };
  if (!expected) fail('Monitor is not set up yet: add a MONITOR_CODE script property in Apps Script.');

  // Slow down guessing: after 10 wrong codes in 10 minutes, refuse for 10 minutes.
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('monitor_fails') || 0);
  if (fails >= 10) fail('Too many wrong access codes. Try again in 10 minutes.');

  if (String(code || '') !== String(expected)) {
    cache.put('monitor_fails', String(fails + 1), 600);
    fail('Wrong access code.');
  }
}

function clearMonitorCache_() {
  try { CacheService.getScriptCache().remove('monitor_v1'); } catch (e) { /* ignore */ }
}

// ---------------------------------------------------------------------------
// Task sheet helpers
// ---------------------------------------------------------------------------

function readTasks_(ss) {
  var sheet = getOrCreateSheet_(ss, TASK_TAB, TASK_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, TASK_HEADERS.length).getValues();
  var tasks = [];
  var fills = [];
  values.forEach(function(r, i) {
    var hasContent = r.some(function(v) { return v !== '' && v !== null; });
    if (!hasContent) return;
    // Rows typed straight into the Sheet by a supervisor: fill in the blanks.
    var changed = false;
    if (!r[TCOL['Task ID']]) { r[TCOL['Task ID']] = newTaskId_(); changed = true; }
    if (!r[TCOL['Status']]) { r[TCOL['Status']] = 'Open'; changed = true; }
    if (!r[TCOL['Source']]) { r[TCOL['Source']] = 'Supervisor'; changed = true; }
    if (!r[TCOL['Created At']]) { r[TCOL['Created At']] = new Date(); changed = true; }
    if (changed) fills.push({ row: i + 2, values: r });
    tasks.push(taskFromRow_(r));
  });
  fills.forEach(function(f) {
    sheet.getRange(f.row, 1, 1, TASK_HEADERS.length).setValues([f.values]);
    formatTaskTimes_(sheet, f.row, 1);
  });
  return tasks;
}

function taskFromRow_(r) {
  return {
    taskId: String(r[TCOL['Task ID']]),
    warehouseId: String(r[TCOL['Warehouse ID']] || ''),
    customerNo: String(r[TCOL['Customer No']] || ''),
    customerName: String(r[TCOL['Customer Name']] || ''),
    locations: String(r[TCOL['Locations']] || ''),
    assignedTo: String(r[TCOL['Assigned To']] || ''),
    status: normStatus_(r[TCOL['Status']]),
    source: String(r[TCOL['Source']] || ''),
    createdAt: toIso_(r[TCOL['Created At']]),
    createdBy: String(r[TCOL['Created By']] || ''),
    startedAt: toIso_(r[TCOL['Started At']]),
    completedAt: toIso_(r[TCOL['Completed At']]),
    expectedPallets: Number(r[TCOL['Expected Pallets']] || 0),
    notes: String(r[TCOL['Notes']] || '')
  };
}

function addTask_(ss, t) {
  var sheet = getOrCreateSheet_(ss, TASK_TAB, TASK_HEADERS);
  var row = [];
  row[TCOL['Task ID']] = newTaskId_();
  row[TCOL['Warehouse ID']] = String(t.warehouseId || '').trim();
  row[TCOL['Customer No']] = String(t.customerNo || '').trim();
  row[TCOL['Customer Name']] = String(t.customerName || '').trim();
  row[TCOL['Locations']] = String(t.locations || '').trim();
  row[TCOL['Assigned To']] = String(t.assignedTo || '').trim();
  row[TCOL['Status']] = t.status || 'Open';
  row[TCOL['Source']] = t.source || '';
  row[TCOL['Created At']] = new Date();
  row[TCOL['Created By']] = t.createdBy || '';
  row[TCOL['Started At']] = sheetTime_(t.startedAt);
  row[TCOL['Completed At']] = '';
  row[TCOL['Expected Pallets']] = '';
  row[TCOL['Notes']] = String(t.notes || '').trim();
  var at = sheet.getLastRow() + 1;
  sheet.getRange(at, 1, 1, TASK_HEADERS.length).setValues([row]);
  formatTaskTimes_(sheet, at, 1);
  return taskFromRow_(row);
}

function findTask_(ss, taskId) {
  if (!taskId) return null;
  var sheet = getOrCreateSheet_(ss, TASK_TAB, TASK_HEADERS);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var values = sheet.getRange(2, 1, lastRow - 1, TASK_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][TCOL['Task ID']]) === String(taskId)) {
      return { sheet: sheet, row: i + 2, task: taskFromRow_(values[i]) };
    }
  }
  return null;
}

function writeTaskFields_(sheet, row, updates) {
  Object.keys(updates).forEach(function(h) {
    var cell = sheet.getRange(row, TCOL[h] + 1);
    cell.setValue(updates[h]);
    if (Object.prototype.toString.call(updates[h]) === '[object Date]') cell.setNumberFormat(DATE_TIME_FORMAT);
  });
}

function newTaskId_() {
  var d = new Date();
  var stamp = Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Los_Angeles', 'yyMMdd-HHmm');
  return 'T' + stamp + '-' + Math.random().toString(36).slice(2, 5).toUpperCase();
}

function normStatus_(s) {
  var v = String(s || '').trim().toLowerCase();
  if (!v || v === 'open' || v === 'new' || v === 'assigned') return 'Open';
  if (v === 'in progress' || v === 'in-progress' || v === 'started' || v === 'counting') return 'In Progress';
  if (v === 'completed' || v === 'complete' || v === 'done') return 'Completed';
  if (v === 'cancelled' || v === 'canceled') return 'Cancelled';
  return String(s).trim();
}

function normWh_(s) { return String(s || '').trim().toUpperCase(); }
function normName_(s) { return String(s || '').trim().toLowerCase(); }

// ---------------------------------------------------------------------------
// One-off cleanup: run from the Apps Script editor (select cleanupBlankAuditRows, Run).
// Removes rows with no Pallet LP and no Record ID from Master_Audit_Log and every
// warehouse tab. Real counts always have a Pallet LP, so they are never touched.
// (An older backend saved monitor/task requests as such blank rows.)
// ---------------------------------------------------------------------------
function cleanupBlankAuditRows() {
  var ss = getSpreadsheet_();
  var skip = [TASK_TAB, EXPECTED_TAB];
  var report = [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (skip.indexOf(name) !== -1 || sheet.getLastRow() < 2) return;
    var header = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    var lpCol = header.indexOf('Pallet LP');
    var idCol = header.indexOf('Record ID');
    if (lpCol === -1) return; // not an audit tab
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, header.length).getValues();
    var removed = 0;
    for (var i = values.length - 1; i >= 0; i--) {
      var noLp = String(values[i][lpCol] || '').trim() === '';
      var noId = idCol === -1 || String(values[i][idCol] || '').trim() === '';
      if (noLp && noId) { sheet.deleteRow(i + 2); removed++; }
    }
    if (removed) report.push(name + ': ' + removed);
    if (name === 'UNASSIGNED' && sheet.getLastRow() < 2 && ss.getSheets().length > 1) {
      ss.deleteSheet(sheet);
      report.push('UNASSIGNED tab removed (empty)');
    }
  });
  clearMonitorCache_();
  var msg = report.length ? 'Removed blank rows: ' + report.join(', ') : 'No blank rows found.';
  Logger.log(msg);
  return msg;
}

// ---------------------------------------------------------------------------
// One-off fix: run from the Apps Script editor (select fixSheetTimes, Run).
// Older versions wrote times as UTC text ("2026-11-23T15:40:12.345Z"), 8 hours ahead
// of California. This turns them into real date-times shown in local time, and sets
// the Sheet's time zone to America/Los_Angeles if it is something else.
// Safe to run more than once: cells that are already date-times are left alone.
// ---------------------------------------------------------------------------
function fixSheetTimes() {
  var ss = getSpreadsheet_();
  var report = [];
  var tz = ss.getSpreadsheetTimeZone();
  if (tz !== SHEET_TIME_ZONE) {
    ss.setSpreadsheetTimeZone(SHEET_TIME_ZONE);
    report.push('Sheet time zone ' + tz + ' -> ' + SHEET_TIME_ZONE);
  }
  ss.getSheets().forEach(function(sheet) {
    var last = sheet.getLastRow();
    if (last < 2) return;
    var head = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    var cols;
    var name = sheet.getName();
    if (name === TASK_TAB) cols = TASK_TIME_COLS;
    else if (name === EXPECTED_TAB) cols = ['Uploaded At'];
    else if (head.indexOf('Pallet LP') !== -1) cols = ['Timestamp'];
    else return;
    var fixed = 0;
    cols.forEach(function(h) {
      var c = head.indexOf(h);
      if (c === -1) return;
      var range = sheet.getRange(2, c + 1, last - 1, 1);
      var values = range.getValues();
      var changed = false;
      values.forEach(function(r) {
        if (typeof r[0] === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(r[0])) {
          var d = toDate_(r[0]);
          if (d) { r[0] = d; changed = true; fixed++; }
        }
      });
      if (changed) range.setValues(values);
      range.setNumberFormat(DATE_TIME_FORMAT);
    });
    if (fixed) report.push(name + ': ' + fixed + ' times converted');
  });
  clearMonitorCache_();
  var msg = report.length ? report.join('; ') : 'All times were already date-times.';
  Logger.log(msg);
  return msg;
}

// ---------------------------------------------------------------------------
// Archive: keeps the last ARCHIVE_KEEP_DAYS days in this Sheet and moves older rows
// to one archive spreadsheet per year ("CycleCount Archive 2026"), one tab per source
// tab and month ("Master_Audit_Log 2026-08"). Moved: count rows of Master_Audit_Log and
// every warehouse tab, plus Completed/Cancelled tasks (Count_Tasks) and their pallet
// lists (Task_Expected). Open tasks and their lists always stay.
// Before anything is removed, the tabs that will change are copied to a backup
// spreadsheet ("CycleCount Backup (before last archive)"), replaced on every run.
// ---------------------------------------------------------------------------
var ARCHIVE_KEEP_DAYS = 90;

// Run once from the editor: creates the monthly trigger (1st of each month, ~2 AM).
function setupMonthlyArchive() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runMonthlyArchive') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runMonthlyArchive').timeBased().onMonthDay(1).atHour(2).create();
  var msg = 'Monthly archive scheduled: 1st of each month around 2 AM (' + Session.getScriptTimeZone() +
    '). Keeps the last ' + ARCHIVE_KEEP_DAYS + ' days.';
  Logger.log(msg);
  return msg;
}

function runMonthlyArchive() {
  var res = archiveOldRecords_({});
  Logger.log(JSON.stringify(res));
  if (!res.ok) throw new Error(res.error); // shows in Executions and the failure email
  return res;
}

// Monitor page: { code, dryRun } — dryRun only reports what would be moved.
function monitorArchive(params) {
  params = params || {};
  try {
    checkMonitorCode_(params.code);
  } catch (err) {
    return { ok: false, error: errorMessage_(err), authError: true };
  }
  return archiveOldRecords_({ dryRun: !!params.dryRun });
}

function archiveOldRecords_(opts) {
  opts = opts || {};
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: 'Sheet is busy (lock timeout). Please retry.' };
  try {
    var ss = getSpreadsheet_();
    var tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    var cutoff = new Date(Date.now() - ARCHIVE_KEEP_DAYS * 86400000);
    var plans = planArchive_(ss, cutoff);
    var total = plans.reduce(function(n, p) { return n + p.old.length; }, 0);
    var summary = { ok: true, dryRun: !!opts.dryRun, keepDays: ARCHIVE_KEEP_DAYS,
      cutoff: cutoff.toISOString(), total: total, moved: {} };
    plans.forEach(function(p) { if (p.old.length) summary.moved[p.name] = p.old.length; });
    if (opts.dryRun || !total) return summary;

    var changing = plans.filter(function(p) { return p.old.length; });
    summary.backupUrl = backupTabs_(changing.map(function(p) { return p.sheet; }), tz);

    // 1) copy old rows to the archive (by year, then source tab + month)
    var archives = {};
    changing.forEach(function(p) {
      var groups = {};
      p.old.forEach(function(o) {
        var key = Utilities.formatDate(o.date, tz, 'yyyy-MM');
        (groups[key] = groups[key] || []).push(o);
      });
      Object.keys(groups).sort().forEach(function(month) {
        var year = month.slice(0, 4);
        var arc = archives[year] = archives[year] || getArchiveSpreadsheet_(year);
        var tab = getOrCreateSheet_(arc, (p.name + ' ' + month).slice(0, 100), p.header);
        var rows = groups[month];
        var range = tab.getRange(tab.getLastRow() + 1, 1, rows.length, p.header.length);
        range.setNumberFormats(rows.map(function(o) { return o.formats; }));
        range.setValues(rows.map(function(o) { return o.values; }));
      });
    });
    SpreadsheetApp.flush();
    summary.archiveUrls = Object.keys(archives).sort().map(function(y) { return archives[y].getUrl(); });

    // 2) only then remove them here, bottom-up in runs of adjacent rows
    changing.forEach(function(p) {
      var rowsIdx = p.old.map(function(o) { return o.row; }).sort(function(a, b) { return b - a; });
      var i = 0;
      while (i < rowsIdx.length) {
        var end = rowsIdx[i], start = end;
        while (i + 1 < rowsIdx.length && rowsIdx[i + 1] === start - 1) { i++; start--; }
        deleteRowsSafe_(p.sheet, start, end - start + 1);
        i++;
      }
    });
    SpreadsheetApp.flush();
    clearMonitorCache_();
    summary.ranAt = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('ARCHIVE_LAST_RUN', JSON.stringify(summary));
    return summary;
  } catch (err) {
    return { ok: false, error: errorMessage_(err) };
  } finally {
    lock.releaseLock();
  }
}

// Works out which rows of which tabs are older than the cutoff. Nothing is changed.
function planArchive_(ss, cutoff) {
  var plans = [];
  var oldTasks = {};
  var taskSheet = ss.getSheetByName(TASK_TAB);
  if (taskSheet) {
    var tp = readTabForArchive_(taskSheet);
    var tId = tp.header.indexOf('Task ID'), tSt = tp.header.indexOf('Status');
    var tDone = tp.header.indexOf('Completed At'), tMade = tp.header.indexOf('Created At');
    tp.old = [];
    if (tId !== -1 && tSt !== -1) {
      tp.rows.forEach(function(r) {
        var st = String(r.values[tSt] || '');
        if (st !== 'Completed' && st !== 'Cancelled') return;
        var d = toDate_(tDone !== -1 && r.values[tDone] ? r.values[tDone] : (tMade !== -1 ? r.values[tMade] : ''));
        if (d && d < cutoff) { r.date = d; tp.old.push(r); oldTasks[String(r.values[tId])] = d; }
      });
    }
    plans.push(tp);
  }
  var expSheet = ss.getSheetByName(EXPECTED_TAB);
  if (expSheet) {
    var ep = readTabForArchive_(expSheet);
    var eId = ep.header.indexOf('Task ID');
    ep.old = eId === -1 ? [] : ep.rows.filter(function(r) {
      var d = oldTasks[String(r.values[eId])];
      if (d) r.date = d;
      return !!d;
    });
    plans.push(ep);
  }
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name === TASK_TAB || name === EXPECTED_TAB || sheet.getLastRow() < 2) return;
    var head = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
    if (head.indexOf('Pallet LP') === -1 || head.indexOf('Timestamp') === -1) return; // not an audit tab
    var p = readTabForArchive_(sheet);
    var tsCol = p.header.indexOf('Timestamp');
    p.old = p.rows.filter(function(r) {
      var d = toDate_(r.values[tsCol]);
      if (d) r.date = d;
      return d && d < cutoff; // rows without a readable date always stay
    });
    plans.push(p);
  });
  plans.forEach(function(p) { delete p.rows; });
  return plans;
}

function readTabForArchive_(sheet) {
  var width = Math.max(1, sheet.getLastColumn());
  var header = sheet.getRange(1, 1, 1, width).getValues()[0];
  var rows = [];
  var last = sheet.getLastRow();
  if (last >= 2) {
    var range = sheet.getRange(2, 1, last - 1, width);
    var values = range.getValues(), formats = range.getNumberFormats();
    values.forEach(function(v, i) {
      if (v.some(function(x) { return x !== '' && x !== null; })) rows.push({ row: i + 2, values: v, formats: formats[i] });
    });
  }
  return { sheet: sheet, name: sheet.getName(), header: header, rows: rows, old: [] };
}

function getArchiveSpreadsheet_(year) {
  var props = PropertiesService.getScriptProperties();
  var key = 'ARCHIVE_ID_' + year;
  var id = props.getProperty(key);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* deleted or no access: make a new one */ }
  }
  var arc = SpreadsheetApp.create('CycleCount Archive ' + year);
  var about = arc.getSheets()[0].setName('About');
  about.getRange(1, 1, 3, 2).setValues([
    ['CycleCount Pro archive', year],
    ['What', 'Count rows and closed tasks older than ' + ARCHIVE_KEEP_DAYS + ' days, moved here from the live Sheet'],
    ['Live Sheet', getSpreadsheet_().getUrl()]
  ]);
  props.setProperty(key, arc.getId());
  return arc;
}

// Copies the given tabs to the backup spreadsheet, replacing the previous backup.
function backupTabs_(sheets, tz) {
  var props = PropertiesService.getScriptProperties();
  var backup = null;
  var id = props.getProperty('ARCHIVE_BACKUP_ID');
  if (id) { try { backup = SpreadsheetApp.openById(id); } catch (e) { backup = null; } }
  if (!backup) {
    backup = SpreadsheetApp.create('CycleCount Backup (before last archive)');
    props.setProperty('ARCHIVE_BACKUP_ID', backup.getId());
  }
  var temp = backup.insertSheet('_tmp_' + Date.now());
  backup.getSheets().forEach(function(sh) { if (sh.getSheetId() !== temp.getSheetId()) backup.deleteSheet(sh); });
  sheets.forEach(function(sh) { sh.copyTo(backup).setName(sh.getName()); });
  var info = backup.insertSheet('Backup info', 0);
  info.getRange(1, 1, 2, 2).setValues([
    ['Backed up at', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')],
    ['From', getSpreadsheet_().getUrl()]
  ]);
  backup.deleteSheet(temp);
  return backup.getUrl();
}

// A sheet must keep at least one row below its frozen header.
function deleteRowsSafe_(sheet, start, count) {
  var spare = sheet.getMaxRows() - Math.max(1, sheet.getFrozenRows()) - count;
  if (spare < 1) sheet.insertRowsAfter(sheet.getMaxRows(), 1 - spare);
  sheet.deleteRows(start, count);
}

function lastArchiveInfo_() {
  var info = { keepDays: ARCHIVE_KEEP_DAYS, lastRun: null, lastTotal: 0 };
  try {
    var last = JSON.parse(PropertiesService.getScriptProperties().getProperty('ARCHIVE_LAST_RUN') || 'null');
    if (last) { info.lastRun = last.ranAt || null; info.lastTotal = last.total || 0; }
  } catch (e) { /* ignore */ }
  return info;
}

function toDate_(v) {
  if (v === '' || v === null || v === undefined) return null;
  var d = Object.prototype.toString.call(v) === '[object Date]' ? v : new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// ---------------------------------------------------------------------------
// General helpers
// ---------------------------------------------------------------------------

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return { ok: false, error: 'Sheet is busy (lock timeout). Please retry.' };
  }
  try {
    return fn();
  } catch (err) {
    return { ok: false, error: errorMessage_(err) };
  } finally {
    lock.releaseLock();
  }
}

function loadTemplate_(names) {
  var lastErr;
  for (var i = 0; i < names.length; i++) {
    try {
      return HtmlService.createTemplateFromFile(names[i]);
    } catch (err) {
      lastErr = err;
    }
  }
  throw new Error('No HTML file named "' + names[0] + '" found in this Apps Script project. ' + errorMessage_(lastErr));
}

function getSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('No spreadsheet linked. Create the script from the Sheet (Extensions -> Apps Script) ' +
    'or set the SPREADSHEET_ID script property.');
}

// Creates the tab with headers, or adds any missing header columns to an existing tab
// (tabs created by older versions keep their data; new columns are added on the right).
function getOrCreateSheet_(ss, tabName, headers) {
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    styleHeader_(sheet, headers.length);
    return sheet;
  }
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  var current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsFix = false;
  for (var i = 0; i < headers.length; i++) {
    if (current[i] !== headers[i]) {
      if (current[i] === '' || current[i] === null) { current[i] = headers[i]; needsFix = true; }
    }
  }
  if (needsFix) {
    sheet.getRange(1, 1, 1, headers.length).setValues([current]);
    styleHeader_(sheet, headers.length);
  }
  return sheet;
}

function styleHeader_(sheet, width) {
  sheet.getRange(1, 1, 1, width)
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
    sheetTime_(r.countedAt) || new Date(),
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
    recordId || '',
    r.taskId || '',
    r.customerNo || '',
    r.customerName || ''
  ];
}

function safeTabName_(s) {
  return String(s).replace(/[^\w-]/g, '_');
}

// ISO text or Date -> Date for the Sheet; blank stays blank; unreadable text is kept as is.
function sheetTime_(v) {
  if (v === '' || v === null || v === undefined) return '';
  var d = toDate_(v);
  return d || v;
}

function formatTaskTimes_(sheet, row, count) {
  TASK_TIME_COLS.forEach(function(h) {
    sheet.getRange(row, TCOL[h] + 1, count, 1).setNumberFormat(DATE_TIME_FORMAT);
  });
}

function toIso_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : v.toISOString();
  var d = new Date(v);
  return isNaN(d.getTime()) ? String(v) : d.toISOString();
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorMessage_(err) {
  return err && err.message ? err.message : String(err);
}
