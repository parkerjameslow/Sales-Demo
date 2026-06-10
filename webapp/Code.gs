/**
 * Demo Request — Form & Dashboard web app.
 *
 * Deploy (one time):
 *   1. script.google.com → New project → add these files:
 *        Code.gs (this file), Form.html, Dashboard.html
 *   2. Deploy → New deployment → type "Web app"
 *        - Execute as: Me
 *        - Who has access: Anyone with the link (or your org)
 *   3. Copy the web app URL:
 *        - Send that URL to reps  → the request form
 *        - Add ?page=dashboard   → the live admin dashboard
 *
 * Submissions land in a Google Sheet ("Demo Request Queue — Live Data"),
 * created automatically on the first submission. Download it as .xlsx any
 * time via File → Download in Sheets.
 */

var SHEET_NAME = 'Requests';
var SPREADSHEET_TITLE = 'Demo Request Queue — Live Data';
var HEADERS = [
  'Request ID', 'Submitted', 'Request Name', 'Requester', 'Email',
  'Application(s)', 'Jira #', 'Sprint #', 'Priority', 'Description', 'Status'
];
var STATUSES = ['New', 'Pending Review', 'Approved', 'Scheduled', 'In Progress', 'Completed', 'Declined'];
var PRIORITIES = ['Critical', 'High', 'Medium', 'Low'];

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'form';
  var file = page === 'dashboard' ? 'Dashboard' : 'Form';
  var template = HtmlService.createTemplateFromFile(file);
  template.baseUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle(page === 'dashboard' ? 'Request Dashboard' : 'Demo Change Request')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Returns the data sheet, creating the spreadsheet on first use. */
function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  var ss;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      ss = null; // sheet was deleted — recreate below
    }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(SPREADSHEET_TITLE);
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#1e1b4b').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Called from Form.html. Validates, appends a row, returns the new ID. */
function submitRequest(data) {
  var required = ['requestName', 'requester', 'applications', 'priority', 'description'];
  for (var i = 0; i < required.length; i++) {
    if (!data || !String(data[required[i]] || '').trim()) {
      throw new Error('Missing required field: ' + required[i]);
    }
  }
  if (PRIORITIES.indexOf(data.priority) === -1) {
    throw new Error('Invalid priority');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getSheet_();
    var nextNum = sheet.getLastRow(); // header row makes the first request #1
    var id = 'REQ-' + ('0000' + nextNum).slice(-4);
    sheet.appendRow([
      id,
      new Date(),
      String(data.requestName).trim(),
      String(data.requester).trim(),
      String(data.email || '').trim(),
      String(data.applications).trim(),
      String(data.jira || '').trim(),
      String(data.sprint || '').trim(),
      data.priority,
      String(data.description).trim(),
      'New'
    ]);
    return { id: id };
  } finally {
    lock.releaseLock();
  }
}

/** Called from Dashboard.html. Returns every request plus config. */
function getDashboardData() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var requests = [];
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      requests.push({
        id: String(r[0]),
        submitted: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
        requestName: String(r[2]),
        requester: String(r[3]),
        email: String(r[4]),
        applications: String(r[5]),
        jira: String(r[6]),
        sprint: String(r[7]),
        priority: String(r[8]),
        description: String(r[9]),
        status: String(r[10] || 'New')
      });
    }
  }
  return {
    requests: requests,
    statuses: STATUSES,
    priorities: PRIORITIES,
    sheetUrl: sheet.getParent().getUrl()
  };
}

/** Called from Dashboard.html when a status dropdown changes. */
function updateStatus(id, status) {
  if (STATUSES.indexOf(status) === -1) throw new Error('Invalid status');
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(id)) {
          sheet.getRange(i + 2, HEADERS.length).setValue(status);
          return { id: id, status: status };
        }
      }
    }
    throw new Error('Request not found: ' + id);
  } finally {
    lock.releaseLock();
  }
}
