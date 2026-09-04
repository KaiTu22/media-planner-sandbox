// Sandbox for the real Project/Version/User schema (§5 of
// End-to-End-System-Plan.md). Read/write client pattern (JSONP reads,
// iframe-POST-then-verify writes) is already proven against the
// closed_deals pilot; this reuses it.
//
// Still not built: Projects/Plans Log derived views (weeks 6-8).

const PARENT_FOLDER_ID = '1tCaf4LoB1qcARLK2faKN2yYTxXMW96lh'; // "Media Planner Projects"

const SHEET_NAMES = {
  users: 'Users',
  projects: 'Projects',
  versions: 'Versions',
  teamRoster: 'TeamRoster',
  agencyHoldCo: 'AgencyHoldCo',
  tentpoleShows: 'TentpoleShows',
};

// slackUserId is optional and manually entered (e.g. copied from a
// teammate's Slack profile "Copy member ID") — not resolved via any Slack
// API. Without it, that person's name appears as plain text in Slack
// notifications rather than a real, notifying mention.
const USER_FIELDS = ['email', 'name', 'role', 'slackUserId'];

// pitchLeadName isn't in the original ER diagram but is required input to
// derive pitchTeam (§5.1) — the current real Assignment sheet has an
// equivalent "Pitch Lead" column feeding the same Maps-tab lookup.
const PROJECT_FIELDS = [
  'id', 'projectName', 'account', 'brand', 'agency', 'holdCo',
  'leadMediaPlannerEmail', 'leadSellerEmail', 'marketingProjectLead',
  'sponsorshipStrategyLead', 'salesAccountManager', 'yieldContact',
  'pitchLeadName', 'pitchTeam', 'rushRequest', 'mediaPlanStatus', 'dealStatus',
  'dealCategory', 'tentpoleShowId', 'folderId', 'driveFolderLink',
  'salesforceLink', 'scratchpadLink', 'budgetSheetLink', 'sponsorshipPlansLink',
  'planRequestDate', 'planDueDate', 'campaignStartDate', 'campaignEndDate',
  'createdAt', 'createdBy', 'updatedAt', 'updatedBy',
  'notifyEmails', // comma-separated emails, picked from the Users list on the Assignment form
];

const VERSION_FIELDS = [
  'id', 'projectId', 'name', 'completedDate', 'totalInvestment',
  'versionStatus', 'folderId', 'packages',
  'createdAt', 'createdBy', 'updatedAt', 'updatedBy',
];

const TEAM_ROSTER_FIELDS = ['teamMemberName', 'pitchTeam'];
const AGENCY_HOLDCO_FIELDS = ['agency', 'holdCo'];
const TENTPOLE_SHOW_FIELDS = ['id', 'name'];

// One-time setup — run manually from the Apps Script editor, not exposed via
// doGet/doPost. Seeds the caller as a write user and a couple of example
// roster rows (from the real Assignment/Logging sheet's Maps tab, §5.1) so
// the lookup derivation has something to resolve against immediately.
// Run this once from the Apps Script editor's "Run" button — clasp's
// headless deploy doesn't trigger the interactive OAuth consent dialog
// needed to grant new scopes (Drive access), only the browser IDE does.
function authorizeDriveAccess() {
  const testValues = { account: 'Debug', brand: 'Debug', projectName: 'Debug (delete me)' };
  createProjectFolder_(testValues);
  Logger.log(testValues.driveFolderLink);
}

// Same reasoning as authorizeDriveAccess — run once from the editor's "Run"
// button after adding MailApp/UrlFetchApp usage, to grant the new scopes.
function authorizeMailAndFetchAccess() {
  MailApp.sendEmail({ to: Session.getActiveUser().getEmail(), subject: 'Media Planner sandbox — authorization test', body: 'If you got this, MailApp is authorized.' });
  Logger.log('Mail sent to ' + Session.getActiveUser().getEmail());
}

function setupSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_NAMES.users, USER_FIELDS);
  ensureSheet_(ss, SHEET_NAMES.projects, PROJECT_FIELDS);
  ensureSheet_(ss, SHEET_NAMES.versions, VERSION_FIELDS);
  const roster = ensureSheet_(ss, SHEET_NAMES.teamRoster, TEAM_ROSTER_FIELDS);
  const agencyHoldCo = ensureSheet_(ss, SHEET_NAMES.agencyHoldCo, AGENCY_HOLDCO_FIELDS);
  ensureSheet_(ss, SHEET_NAMES.tentpoleShows, TENTPOLE_SHOW_FIELDS);

  const usersSheet = ss.getSheetByName(SHEET_NAMES.users);
  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow([Session.getActiveUser().getEmail(), 'Admin', 'write', '']);
  }
  if (roster.getLastRow() < 2) {
    roster.appendRow(['Nicole Rosenberg', 'ROSENBERG']);
    roster.appendRow(['Bari Zibrak', 'ZIBRAK']);
  }
  if (agencyHoldCo.getLastRow() < 2) {
    agencyHoldCo.appendRow(['Agency D7', 'PMX']);
    agencyHoldCo.appendRow(['Carat', 'Dentsu']);
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

// Column order comes from the sheet's actual header row, not the `fields`
// array's order — appendRow/setValue are purely positional, so if PROJECT_FIELDS
// (etc.) grows over time, writing by array order silently misaligns every
// existing column after the insertion point. This adds any new fields as
// extra columns at the end instead, so growing the schema never shifts
// existing data (confirmed bug, fixed 2026-09-01).
function ensureColumns_(sheet, fields) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const missing = fields.filter(function (f) { return headers.indexOf(f) === -1; });
  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
  return headers.concat(missing);
}

function appendRecord_(sheetName, fields, values) {
  const sheet = getSheet_(sheetName);
  const headers = ensureColumns_(sheet, fields);
  const row = headers.map(function (h) {
    let v = values[h];
    if (h === 'packages' && v && typeof v === 'object') v = JSON.stringify(v);
    return v === undefined || v === null ? '' : v;
  });
  sheet.appendRow(row);
}

function updateRecord_(sheetName, fields, idField, id, values) {
  const sheet = getSheet_(sheetName);
  const headers = ensureColumns_(sheet, fields);
  const data = sheet.getDataRange().getValues();
  const idIdx = headers.indexOf(idField);
  for (let r = 1; r < data.length; r++) {
    if (data[r][idIdx] === id) {
      fields.forEach(function (f) {
        if (Object.prototype.hasOwnProperty.call(values, f)) {
          let v = values[f];
          if (f === 'packages' && v && typeof v === 'object') v = JSON.stringify(v);
          const colIdx = headers.indexOf(f);
          sheet.getRange(r + 1, colIdx + 1).setValue(v);
        }
      });
      return true;
    }
  }
  return false;
}

// §5.1 — pitchTeam and holdCo are lookup-derived, never manually typed.
function lookupValue_(sheetName, keyField, key, valueField) {
  if (!key) return null;
  const rows = readRows_(sheetName);
  const match = rows.find(function (r) { return r[keyField] === key; });
  return match ? match[valueField] : null;
}

function applyProjectLookups_(values) {
  if (values.pitchLeadName) {
    values.pitchTeam = lookupValue_(SHEET_NAMES.teamRoster, 'teamMemberName', values.pitchLeadName, 'pitchTeam');
  }
  if (values.agency) {
    values.holdCo = lookupValue_(SHEET_NAMES.agencyHoldCo, 'agency', values.agency, 'holdCo');
  }
}

// §5.2 — auto-creates a subfolder under the shared parent, named
// "{Account} - {Brand} - {ProjectName}" (confirmed 2026-08-31).
function createProjectFolder_(values) {
  const name = [values.account, values.brand, values.projectName].filter(Boolean).join(' - ');
  const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
  const folder = parent.createFolder(name || 'Untitled Project');
  values.folderId = folder.getId();
  values.driveFolderLink = folder.getUrl();
}

// General-purpose file upload (confirmed 2026-09-04) — lets a planner
// attach any supporting document (signed IO, client PDF, screenshot, etc.)
// to a project's Drive folder from inside the app, instead of clicking
// through to Drive and finding the folder themselves. Content arrives
// base64-encoded (safe for a form POST body, unlike raw binary).
function uploadProjectAttachment_(projectId, filename, mimeType, contentBase64) {
  const project = readRows_(SHEET_NAMES.projects).find(function (p) { return p.id === projectId; });
  if (!project || !project.folderId) {
    throw new Error('Project has no Drive folder to upload to.');
  }
  const folder = DriveApp.getFolderById(project.folderId);
  const blob = Utilities.newBlob(Utilities.base64Decode(contentBase64), mimeType || 'application/octet-stream', filename);
  const file = folder.createFile(blob);
  return { id: file.getId(), name: file.getName(), url: file.getUrl() };
}

function listProjectFiles_(projectId) {
  const project = readRows_(SHEET_NAMES.projects).find(function (p) { return p.id === projectId; });
  if (!project || !project.folderId) return [];
  const folder = DriveApp.getFolderById(project.folderId);
  const files = folder.getFiles();
  const result = [];
  while (files.hasNext()) {
    const f = files.next();
    result.push({ id: f.getId(), name: f.getName(), url: f.getUrl(), lastUpdated: f.getLastUpdated().toISOString() });
  }
  return result;
}

// Automatically keeps one continuously-updated INTERNAL_*.json file in the
// project's Drive folder — replaces the old manual "Export Internal Plan"
// download-then-drag-into-Drive workflow (confirmed 2026-09-04). Overwrites
// the existing file by name rather than creating a new dated one each time,
// since this now fires automatically on every debounced sync, not on a
// deliberate user click.
function uploadProjectFile_(projectId, content) {
  const project = readRows_(SHEET_NAMES.projects).find(function (p) { return p.id === projectId; });
  if (!project || !project.folderId) return; // no Drive folder yet (e.g. legacy/imported project) — skip
  const sanitizedName = (project.projectName || 'untitled').replace(/[^a-z0-9]/gi, '_').toLowerCase();
  const filename = 'INTERNAL_' + sanitizedName + '.json';
  const folder = DriveApp.getFolderById(project.folderId);
  const existing = folder.getFilesByName(filename);
  if (existing.hasNext()) {
    existing.next().setContent(content);
  } else {
    folder.createFile(filename, content, MimeType.PLAIN_TEXT);
  }
}

// §6.1 step 5 — notifies the assigned Lead Media Planner plus anyone else
// picked on the Assignment form. Email works today via MailApp (no external
// service, generous Workspace quota). Slack posts to one pre-existing
// channel via an Incoming Webhook (Script Property SLACK_WEBHOOK_URL) — set
// that property in the Apps Script editor's Project Settings once a webhook
// exists; until then this just logs and skips, same as a missing recipient.
// Never blocks project creation — both paths are wrapped by the caller.
function sendAssignmentNotifications_(project) {
  const emails = [project.leadMediaPlannerEmail]
    .concat((project.notifyEmails || '').split(','))
    .map(function (e) { return (e || '').trim(); })
    .filter(Boolean);
  if (emails.length === 0) return;

  const dedupedEmails = Array.from(new Set(emails));
  const subject = 'New Assignment: ' + (project.projectName || 'Untitled Project');
  const lines = [
    'Account / Brand: ' + [project.account, project.brand].filter(Boolean).join(' / '),
    project.agency ? 'Agency: ' + project.agency : null,
    project.planDueDate ? 'Plan Due: ' + project.planDueDate : null,
    project.rushRequest ? 'RUSH REQUEST' : null,
    project.driveFolderLink ? 'Drive folder: ' + project.driveFolderLink : null,
  ].filter(Boolean);
  const body = lines.join('\n');

  try {
    MailApp.sendEmail({ to: dedupedEmails.join(','), subject: subject, body: body });
  } catch (e) {
    Logger.log('Email notification failed: ' + e.message);
  }

  try {
    sendSlackNotification_(subject, lines, dedupedEmails);
  } catch (e) {
    Logger.log('Slack notification failed: ' + e.message);
  }
}

function sendSlackNotification_(subject, lines, emails) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhookUrl) {
    Logger.log('Slack not configured (no SLACK_WEBHOOK_URL script property) — skipping.');
    return;
  }
  const users = readRows_(SHEET_NAMES.users);
  const mentions = emails.map(function (email) {
    const user = users.find(function (u) { return u.email === email; });
    if (user && user.slackUserId) return '<@' + user.slackUserId + '>';
    return (user && user.name) || email;
  });
  const text = '*' + subject + '*\n' + lines.join('\n') + '\nNotifying: ' + mentions.join(', ');
  UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
  });
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
    if (action === 'listProjectFiles') {
      result = listProjectFiles_(e.parameter.projectId);
    } else if (action === 'listProjects') {
      result = readRows_(SHEET_NAMES.projects);
    } else if (action === 'listVersions') {
      result = readRows_(SHEET_NAMES.versions);
    } else if (action === 'listUsers') {
      result = readRows_(SHEET_NAMES.users);
    } else if (action === 'listTeamRoster') {
      result = readRows_(SHEET_NAMES.teamRoster);
    } else if (action === 'listAgencyHoldCo') {
      result = readRows_(SHEET_NAMES.agencyHoldCo);
    } else if (action === 'listTentpoleShows') {
      result = readRows_(SHEET_NAMES.tentpoleShows);
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

// Holds the script lock only around the actual sheet mutation, not around
// slow unrelated I/O (Drive folder creation, lookups, notifications).
// Confirmed 2026-09-04 by a real concurrency test: holding the lock across
// createProject's full body (including Drive folder creation) meant a
// queue of concurrent writes serialized through that slow step too, and
// 2 of 10 concurrent test writes timed out waiting for the lock and were
// silently lost (waitLock() throws outside try/finally, and doPost's
// response is unreadable anyway — see comment below). Narrowing the lock
// to just the sheet write keeps hold time to milliseconds, so far more
// concurrent writers can be served within the same timeout.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    fn();
  } finally {
    lock.releaseLock();
  }
}

// Response body is unreadable cross-origin (script.google.com sends
// X-Frame-Options: sameorigin, blocking any iframe response including
// postMessage — confirmed against the closed_deals pilot). Callers must
// verify writes via a follow-up doGet read, not this response.
function doPost(e) {
  const action = e.parameter.action;
  const user = getCurrentUser_();
  const values = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
  const now = new Date().toISOString();

  if (action === 'uploadProjectAttachment') {
    requireWrite_(user);
    uploadProjectAttachment_(e.parameter.projectId, e.parameter.filename, e.parameter.mimeType, e.parameter.contentBase64);
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  } else if (action === 'uploadProjectFile') {
    requireWrite_(user);
    uploadProjectFile_(e.parameter.projectId, e.parameter.content);
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  } else if (action === 'createProject') {
    requireWrite_(user);
    values.id = values.id || Utilities.getUuid();
    // §6.1 step 4 — Media Plan Status defaults to Pre-Planning; Deal
    // Status stays unset until the deal actually resolves (§6.2 step 5).
    values.mediaPlanStatus = values.mediaPlanStatus || 'Pre-Planning';
    values.createdAt = now;
    values.createdBy = user.email;
    values.updatedAt = now;
    values.updatedBy = user.email;
    applyProjectLookups_(values); // reads TeamRoster/AgencyHoldCo — different sheets, no lock needed
    createProjectFolder_(values); // Drive API call — slow, no shared-sheet state, no lock needed
    withLock_(function () {
      appendRecord_(SHEET_NAMES.projects, PROJECT_FIELDS, values);
    });
    try {
      sendAssignmentNotifications_(values);
    } catch (e) {
      Logger.log('Assignment notifications failed: ' + e.message);
    }
  } else if (action === 'updateProject') {
    requireWrite_(user);
    values.updatedAt = now;
    values.updatedBy = user.email;
    applyProjectLookups_(values);
    withLock_(function () {
      updateRecord_(SHEET_NAMES.projects, PROJECT_FIELDS, 'id', values.id, values);
    });
  } else if (action === 'createVersion') {
    requireWrite_(user);
    values.id = values.id || Utilities.getUuid();
    values.createdAt = now;
    values.createdBy = user.email;
    values.updatedAt = now;
    values.updatedBy = user.email;
    withLock_(function () {
      appendRecord_(SHEET_NAMES.versions, VERSION_FIELDS, values);
    });
  } else if (action === 'updateVersion') {
    requireWrite_(user);
    values.updatedAt = now;
    values.updatedBy = user.email;
    withLock_(function () {
      updateRecord_(SHEET_NAMES.versions, VERSION_FIELDS, 'id', values.id, values);
    });
  } else if (action === 'createTeamRosterEntry') {
    requireWrite_(user);
    withLock_(function () {
      appendRecord_(SHEET_NAMES.teamRoster, TEAM_ROSTER_FIELDS, values);
    });
  } else if (action === 'createAgencyHoldCoEntry') {
    requireWrite_(user);
    withLock_(function () {
      appendRecord_(SHEET_NAMES.agencyHoldCo, AGENCY_HOLDCO_FIELDS, values);
    });
  } else if (action === 'createTentpoleShow') {
    requireWrite_(user);
    values.id = values.id || Utilities.getUuid();
    withLock_(function () {
      appendRecord_(SHEET_NAMES.tentpoleShows, TENTPOLE_SHOW_FIELDS, values);
    });
  }

  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
