/**
 * Admin Queue dashboard — Google Sheets generator
 * ------------------------------------------------
 * Paste into https://script.google.com (new project), then run
 * buildAdminQueue(). It creates a new Google Sheet with:
 *   - 5 KPI cards (live COUNTIF formulas)
 *   - a request table with Priority/Status dropdowns, color-coded priority,
 *     colored status pills, a frozen header and a filter
 * The Sheet URL is printed in the Execution log.
 */

function buildAdminQueue() {
  var ss = SpreadsheetApp.create('Admin Queue');
  var sh = ss.getActiveSheet();
  sh.setName('Admin Queue');
  sh.setHiddenGridlines(true);

  // palette
  var INK='#1F2A44', SLATE='#6B7A90', HAIR='#D8DEE9', WHITE='#FFFFFF', HEADERBG='#F2F5FA';
  var BLUE='#2F5BEA', ORANGE='#E0892B', TEAL='#3BA9C4', GREEN='#3FA34D', RED='#D64545',
      PURPLE='#7C4DD6';

  // column widths (px)
  var widths = [200, 90, 150, 100, 130, 110, 110, 90, 220];
  for (var i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);

  // ---- title ----
  sh.getRange('A1:I1').merge().setValue('Admin Queue')
    .setFontSize(20).setFontWeight('bold').setFontColor(INK);
  sh.getRange('A2:I2').merge()
    .setValue('Requests waiting for your review. Set a Status to move a request through the pipeline.')
    .setFontSize(10).setFontColor(SLATE);

  // ---- KPI cards (rows 4-7, column pairs; last card uses col I alone) ----
  var cards = [
    {c0:1,c1:2, label:'PENDING REVIEW', sub:'Need your decision',     accent:BLUE,
     f:'=COUNTIF(E:E,"New")+COUNTIF(E:E,"In Review")'},
    {c0:3,c1:4, label:'IN PROGRESS',    sub:'Being prepared',         accent:ORANGE,
     f:'=COUNTIF(E:E,"In Progress")'},
    {c0:5,c1:6, label:'SCHEDULED',      sub:'On the calendar',        accent:TEAL,
     f:'=COUNTIF(E:E,"Scheduled")'},
    {c0:7,c1:8, label:'COMPLETED',      sub:'Delivered',              accent:GREEN,
     f:'=COUNTIF(E:E,"Completed")'},
    {c0:9,c1:9, label:'CRITICAL',       sub:'Needs immediate action', accent:RED,
     f:'=COUNTIF(D:D,"Critical")'}
  ];
  sh.setRowHeight(4, 6).setRowHeight(5, 40).setRowHeight(6, 18).setRowHeight(7, 16);
  cards.forEach(function (k) {
    var n = k.c1 - k.c0 + 1;
    sh.getRange(4, k.c0, 1, n).merge().setBackground(k.accent);            // accent bar
    sh.getRange(5, k.c0, 1, n).merge().setFormula(k.f)
      .setFontSize(26).setFontWeight('bold').setFontColor(k.accent)
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setBackground(WHITE);
    sh.getRange(6, k.c0, 1, n).merge().setValue(k.label)
      .setFontSize(9).setFontWeight('bold').setFontColor(SLATE).setBackground(WHITE);
    sh.getRange(7, k.c0, 1, n).merge().setValue(k.sub)
      .setFontSize(8).setFontColor(SLATE).setBackground(WHITE);
    sh.getRange(4, k.c0, 4, n).setBorder(true, true, true, true, false, false, HAIR,
      SpreadsheetApp.BorderStyle.SOLID);
  });

  // ---- table ----
  var HEADER_ROW = 10;
  var headers = ['Request','Requester','Application(s)','Priority','Status',
                 'Jira #','Sprint','Submitted','Notes'];
  sh.getRange(HEADER_ROW, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setFontColor(INK).setBackground(HEADERBG);

  var rows = [
    ['Agent — Screen Pop Customization',     'Sarah M.',  'Agent',              'High',     'Scheduled',   'DEMO-482','Sprint 44','May 2',''],
    ['WFM — Real-Time Adherence Dashboard',   'James T.',  'WFM, Supervisor',    'Critical', 'In Progress', 'DEMO-479','Sprint 44','Apr 30',''],
    ['QM — AI Evaluation Scorecard',          'Priya K.',  'Quality Management', 'Medium',   'Approved',    'DEMO-491','Sprint 45','May 4',''],
    ['Supervisor — Live Monitoring Panel',    'Marcus L.', 'Supervisor',         'High',     'In Review',   '—','—','May 5',''],
    ['Interactions Hub — Recording Playback', 'Dana W.',   'Interactions Hub',   'Low',      'New',         '—','—','May 6',''],
    ['Smart Reach — Campaign Builder Flow',   'Jordan B.', 'Smart Reach',        'High',     'Completed',   'DEMO-465','Sprint 43','Apr 22',''],
    ['My Zone — Schedule Visibility',         'Sarah M.',  'My Zone',            'Medium',   'New',         '—','—','May 6',''],
    ['Agent — After Call Work Timer',         'Priya K.',  'Agent',              'Medium',   'New',         '—','—','May 6','']
  ];
  var first = HEADER_ROW + 1;
  sh.getRange(first, 1, rows.length, headers.length).setValues(rows);
  sh.getRange(first, 1, rows.length, 1).setFontWeight('bold').setFontColor(INK);
  sh.getRange(first, 4, rows.length, 2).setHorizontalAlignment('center').setFontWeight('bold');

  var LAST = first + 60; // room to grow
  sh.getRange(HEADER_ROW, 1, LAST - HEADER_ROW + 1, headers.length)
    .setBorder(true, true, true, true, true, true, HAIR, SpreadsheetApp.BorderStyle.SOLID);

  // dropdowns
  var prio = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Critical','High','Medium','Low'], true).setAllowInvalid(false).build();
  var stat = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New','In Review','Approved','Scheduled','In Progress','Completed','Blocked','Rejected'], true)
    .setAllowInvalid(false).build();
  sh.getRange(first, 4, LAST - first + 1, 1).setDataValidation(prio);
  sh.getRange(first, 5, LAST - first + 1, 1).setDataValidation(stat);

  // conditional formatting — priority + status colors
  var prioRange = sh.getRange(first, 4, LAST - first + 1, 1);
  var statRange = sh.getRange(first, 5, LAST - first + 1, 1);
  var rules = [];
  function cf(range, text, bg, fg) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text).setBackground(bg).setFontColor(fg || INK)
      .setRanges([range]).build());
  }
  cf(prioRange,'Critical','#F7D7D7',RED);
  cf(prioRange,'High','#FCE6CE',ORANGE);
  cf(prioRange,'Medium','#FFF1C9','#9A7B0A');
  cf(prioRange,'Low','#DCEFDD',GREEN);
  cf(statRange,'New','#EDF0F5',SLATE);
  cf(statRange,'In Review','#E3ECFF',BLUE);
  cf(statRange,'Approved','#EEE6FB',PURPLE);
  cf(statRange,'Scheduled','#DFF3F7',TEAL);
  cf(statRange,'In Progress','#FCEBD6',ORANGE);
  cf(statRange,'Completed','#DCEFDD',GREEN);
  cf(statRange,'Blocked','#F7DDDD',RED);
  cf(statRange,'Rejected','#E6E9EE',INK);
  sh.setConditionalFormatRules(rules);

  // freeze + filter
  sh.setFrozenRows(HEADER_ROW);
  sh.getRange(HEADER_ROW, 1, LAST - HEADER_ROW + 1, headers.length).createFilter();

  Logger.log('✅ Admin Queue created!');
  Logger.log('Open it here: ' + ss.getUrl());
}
