// Sandbox skeleton for the real Project/Version/User schema (§5 of
// End-to-End-System-Plan.md). Intentionally basic CRUD only — lookup-table
// derivation (pitchTeam/holdCo), Drive folder auto-creation, and the
// Projects/Plans Log derived views come in later phases (§9, weeks 3-8).
// Read/write client pattern (JSONP reads, iframe-POST-then-verify writes)
// is already proven against the closed_deals pilot; this reuses it.

const SHEET_NAMES = {
  users: 'Users',
  projects: 'Projects',
  versions: 'Versions',
};

const USER_FIELDS = ['email', 'name', 'role'];

const PROJECT_FIELDS = [
  'id', 'projectName', 'account', 'brand', 'agency', 'holdCo',
  'leadMediaPlannerEmail', 'leadSellerEmail', 'marketingProjectLead',
  'sponsorshipStrategyLead', 'salesAccountManager', 'yieldContact',
  'pitchTeam', 'rushRequest', 'mediaPlanStatus', 'dealStatus', 'dealCategory',
  'tentpoleShowId', 'folderId', 'driveFolderLink', 'salesforceLink',
  'scratchpadLink', 'budgetSheetLink', 'sponsorshipPlansLink',
  'planRequestDate', 'planDueDate', 'campaignStartDate', 'campaignEndDate',
  'createdAt', 'createdBy', 'updatedAt', 'updatedBy',
];

const VERSION_FIELDS = [
  'id', 'projectId', 'name', 'completedDate', 'totalInvestment',
  'versionStatus', 'folderId', 'packages',
  'createdAt', 'createdBy', 'updatedAt', 'updatedBy',
];

// One-time setup — run manually from the Apps Script editor, not exposed via
// doGet/doPost. Seeds the caller as a write user so there's an initial admin.
function setupSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_NAMES.users, USER_FIELDS);
  ensureSheet_(ss, SHEET_NAMES.projects, PROJECT_FIELDS);
  ensureSheet_(ss, SHEET_NAMES.versions, VERSION_FIELDS);

  const usersSheet = ss.getSheetByName(SHEET_NAMES.users);
  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow([Session.getActiveUser().getEmail(), 'Admin', 'write']);
  }

  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readRows_(sheetName) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(function (row) { return rowToObject_(headers, row); });
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach(function (header, i) {
    let value = row[i];
    if (header === 'packages' && typeof value === 'string' && value) {
      try {
        value = JSON.parse(value);
      } catch (err) {
        value = null;
      }
    } else if (value instanceof Date) {
      value = value.toISOString();
    } else if (value === '') {
      value = null;
    }
    obj[header] = value;
  });
  return obj;
}

function getCurrentUser_() {
  const email = Session.getActiveUser().getEmail();
  const users = readRows_(SHEET_NAMES.users);
  const record = users.find(function (u) { return u.email === email; });
  return {
    email: email,
    role: record ? record.role : 'read',
    name: record ? record.name : null,
  };
}

function requireWrite_(user) {
  if (user.role !== 'write') {
    throw new Error('Write access required for ' + user.email);
  }
}

function appendRecord_(sheetName, fields, values) {
  const sheet = getSheet_(sheetName);
  const row = fields.map(function (f) {
    let v = values[f];
    if (f === 'packages' && v && typeof v === 'object') v = JSON.stringify(v);
    return v === undefined || v === null ? '' : v;
  });
  sheet.appendRow(row);
}

function updateRecord_(sheetName, fields, idField, id, values) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIdx = headers.indexOf(idField);
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === id) {
      fields.forEach(function (f, i) {
        if (Object.prototype.hasOwnProperty.call(values, f)) {
          let v = values[f];
          if (f === 'packages' && v && typeof v === 'object') v = JSON.stringify(v);
          sheet.getRange(r + 1, i + 1).setValue(v);
        }
      });
      return true;
    }
  }
  return false;
}

function respond_(result, callback) {
  const body = JSON.stringify(result);
  return callback
    ? ContentService.createTextOutput(callback + '(' + body + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result;
  try {
    if (action === 'listProjects') {
      result = readRows_(SHEET_NAMES.projects);
    } else if (action === 'listVersions') {
      result = readRows_(SHEET_NAMES.versions);
    } else if (action === 'listUsers') {
      result = readRows_(SHEET_NAMES.users);
    } else if (action === 'whoami') {
      result = getCurrentUser_();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return respond_(result, callback);
}

// Response body is unreadable cross-origin (script.google.com sends
// X-Frame-Options: sameorigin, blocking any iframe response including
// postMessage — confirmed against the closed_deals pilot). Callers must
// verify writes via a follow-up doGet read, not this response.
function doPost(e) {
  const action = e.parameter.action;
  const user = getCurrentUser_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const values = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    const now = new Date().toISOString();

    if (action === 'createProject') {
      requireWrite_(user);
      values.id = values.id || Utilities.getUuid();
      values.createdAt = now;
      values.createdBy = user.email;
      values.updatedAt = now;
      values.updatedBy = user.email;
      appendRecord_(SHEET_NAMES.projects, PROJECT_FIELDS, values);
    } else if (action === 'updateProject') {
      requireWrite_(user);
      values.updatedAt = now;
      values.updatedBy = user.email;
      updateRecord_(SHEET_NAMES.projects, PROJECT_FIELDS, 'id', values.id, values);
    } else if (action === 'createVersion') {
      requireWrite_(user);
      values.id = values.id || Utilities.getUuid();
      values.createdAt = now;
      values.createdBy = user.email;
      values.updatedAt = now;
      values.updatedBy = user.email;
      appendRecord_(SHEET_NAMES.versions, VERSION_FIELDS, values);
    } else if (action === 'updateVersion') {
      requireWrite_(user);
      values.updatedAt = now;
      values.updatedBy = user.email;
      updateRecord_(SHEET_NAMES.versions, VERSION_FIELDS, 'id', values.id, values);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
