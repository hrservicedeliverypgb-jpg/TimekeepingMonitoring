// Code.gs — Overtime Monitoring REST API
// Deploy: Extensions > Apps Script > Deploy > New deployment
// Type: Web App | Execute as: Me | Who has access: Anyone (or your org)
// After deploying, copy the Web App URL into the GAS_URL constant in index.html

const SHEET_IMPORTS   = 'OT_Imports';
const SHEET_EMPLOYEES = 'OT_Employees';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';
  try {
    if (action === 'getAll') return respond(actionGetAll());
    return respond({ error: 'Unknown GET action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body    = JSON.parse(e.postData.contents);
    const action  = body.action;
    const payload = body.payload;
    if (action === 'save')     return respond(actionSave(payload));
    if (action === 'checkDup') return respond(actionCheckDup(payload));
    if (action === 'delete')   return respond(actionDelete(payload));
    if (action === 'getAll')   return respond(actionGetAll());
    return respond({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// ── Helpers ────────────────────────────────────────────────────────────
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      const range = sheet.getRange(1, 1, 1, headers.length);
      range.setValues([headers]);
      range.setFontWeight('bold').setBackground('#0d1f3c').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getImportsSheet() {
  return getOrCreateSheet(SHEET_IMPORTS, [
    'id', 'affiliate', 'periodLabel', 'periodStart', 'periodEnd',
    'fileName', 'uploadedAt', 'employeeCount',
    'totalRegHours', 'totalOTHours', 'totalHours'
  ]);
}

function getEmployeesSheet() {
  return getOrCreateSheet(SHEET_EMPLOYEES, [
    'importId', 'affiliate', 'periodLabel',
    'name', 'department', 'costCenter',
    'regularHours', 'overtimeHours', 'totalHours', 'otPercentage'
  ]);
}

// ── Actions ────────────────────────────────────────────────────────────
function actionGetAll() {
  const impSheet = getImportsSheet();
  const empSheet = getEmployeesSheet();
  const impData  = impSheet.getDataRange().getValues();
  const empData  = empSheet.getDataRange().getValues();
  if (impData.length <= 1) return { imports: [] };

  const impHeaders = impData[0];
  const empHeaders = empData[0];

  const empByImport = {};
  for (let r = 1; r < empData.length; r++) {
    const obj = {};
    empHeaders.forEach((h, i) => { obj[h] = empData[r][i]; });
    const id = String(obj.importId);
    if (!empByImport[id]) empByImport[id] = [];
    empByImport[id].push({
      name:          obj.name,
      department:    obj.department,
      costCenter:    obj.costCenter,
      regularHours:  Number(obj.regularHours)  || 0,
      overtimeHours: Number(obj.overtimeHours) || 0,
      totalHours:    Number(obj.totalHours)    || 0,
      otPercentage:  Number(obj.otPercentage)  || 0
    });
  }

  const imports = [];
  for (let r = 1; r < impData.length; r++) {
    const obj = {};
    impHeaders.forEach((h, i) => { obj[h] = impData[r][i]; });
    const id = String(obj.id);
    imports.push({
      id:            obj.id,
      affiliate:     obj.affiliate,
      periodLabel:   obj.periodLabel,
      periodStart:   obj.periodStart,
      periodEnd:     obj.periodEnd,
      fileName:      obj.fileName,
      uploadedAt:    obj.uploadedAt,
      employeeCount: Number(obj.employeeCount)  || 0,
      totalRegHours: Number(obj.totalRegHours)  || 0,
      totalOTHours:  Number(obj.totalOTHours)   || 0,
      totalHours:    Number(obj.totalHours)     || 0,
      employees:     empByImport[id] || []
    });
  }
  return { imports };
}

function actionCheckDup(payload) {
  const { affiliate, periodLabel } = payload;
  const sheet   = getImportsSheet();
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return { duplicate: false };
  const headers   = data[0];
  const affIdx    = headers.indexOf('affiliate');
  const periodIdx = headers.indexOf('periodLabel');
  for (let r = 1; r < data.length; r++) {
    if (data[r][affIdx] === affiliate && data[r][periodIdx] === periodLabel) {
      return { duplicate: true, row: r + 1 };
    }
  }
  return { duplicate: false };
}

function actionSave(payload) {
  const impSheet = getImportsSheet();
  const empSheet = getEmployeesSheet();
  const id       = String(payload.id);
  _deleteRowsById(impSheet, 'id', id);
  _deleteRowsById(empSheet, 'importId', id);

  impSheet.appendRow([
    id, payload.affiliate, payload.periodLabel,
    payload.periodStart || '', payload.periodEnd || '',
    payload.fileName, payload.uploadedAt, payload.employeeCount,
    payload.totalRegHours, payload.totalOTHours, payload.totalHours
  ]);

  (payload.employees || []).forEach(e => {
    empSheet.appendRow([
      id, payload.affiliate, payload.periodLabel,
      e.name, e.department || '', e.costCenter || '',
      e.regularHours, e.overtimeHours, e.totalHours, e.otPercentage
    ]);
  });

  return { ok: true, saved: id };
}

function actionDelete(payload) {
  const id = String(payload.id);
  _deleteRowsById(getImportsSheet(),  'id',       id);
  _deleteRowsById(getEmployeesSheet(),'importId', id);
  return { ok: true, deleted: id };
}

function _deleteRowsById(sheet, colName, id) {
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIdx  = headers.indexOf(colName);
  if (colIdx === -1) return;
  for (let r = data.length - 1; r >= 1; r--) {
    if (String(data[r][colIdx]) === id) sheet.deleteRow(r + 1);
  }
}
