/**
 * ============================================================
 *  CompProg 101 — Grade Recorder
 *  Google Apps Script (Web App)
 *  Instructor: DEBULGADO, S.M.G
 * ============================================================
 * SETUP:
 * 1. Extensions → Apps Script → paste this → Save
 * 2. Run → setupSpreadsheet (once)
 * 3. Deploy → Web App → Anyone → Copy URL
 * 4. Paste URL in practical.html and quiz.html
 * ============================================================
 */

// ── SHEET NAMES ──
const SHEET_PRACTICAL = 'practical_cfp_sheet';
const SHEET_THEORY    = 'Theory Quiz Results';
const SHEET_SUMMARY   = 'Student Summary';
const SHEET_LOG       = 'Activity Log';

// ── PRACTICAL COLUMNS (A–P) ──
const PRACTICAL_HEADERS = [
  'No.','Timestamp (PHT)','Full Name','Student Number','Section',
  'Task 1 Result','Task 1 Points','Task 2 Result','Task 2 Points',
  'Total Score','Percentage','Grade','Time Taken','Submission Type',
  'Task 1 Output','Task 2 Output'
];

// ── THEORY COLUMNS ──
const THEORY_HEADERS = [
  'No.','Timestamp (PHT)','Full Name','Student Number','Section',
  'Score','Out Of','Percentage','Grade','Correct','Wrong / Skipped',
  'Time Taken','Answer Details'
];

// ─────────────────────────────────────────
//  MAIN ENTRY
// ─────────────────────────────────────────
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var ts   = new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });

    if (data.type === 'PRACTICAL') {
      if (checkPracticalDuplicate(ss, data)) {
        logActivity(ss, { type:'DUPLICATE_BLOCKED', name:data.name, section:data.section, percent:data.percent, autoSubmit:false }, ts);
        return ContentService
          .createTextOutput(JSON.stringify({ status: 'duplicate', message: 'Already submitted. Only 1 take allowed.' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      handlePractical(ss, data, ts);
    } else {
      handleTheory(ss, data, ts);
    }

    logActivity(ss, data, ts);
    updateSummary(ss, data, ts);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Recorded!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('CompProg 101 Grade Recorder — DEBULGADO, S.M.G — OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Duplicate check: same student number in Practical sheet ──
function checkPracticalDuplicate(ss, data) {
  var sheet = ss.getSheetByName(SHEET_PRACTICAL);
  if (!sheet || sheet.getLastRow() <= 1) return false;
  var incoming = (data.studentNo || '').toString().trim().toLowerCase();
  if (!incoming || incoming === 'n/a') return false;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if ((values[i][3] || '').toString().trim().toLowerCase() === incoming) return true;
  }
  return false;
}

// ─────────────────────────────────────────
//  PRACTICAL HANDLER → writes to "practical_cfp_sheet"
// ─────────────────────────────────────────
function handlePractical(ss, data, ts) {
  var sheet = ss.getSheetByName(SHEET_PRACTICAL);
  if (!sheet) { sheet = ss.insertSheet(SHEET_PRACTICAL); setupPracticalSheet(sheet); }

  var details  = data.details || '';
  var t1Pass   = details.indexOf('Task1:PASS') >= 0;
  var t2Pass   = details.indexOf('Task2:PASS') >= 0;
  var t1Pts    = t1Pass ? 50 : 0;
  var t2Pts    = t2Pass ? 50 : 0;
  var total    = t1Pts + t2Pts;
  var grade    = total >= 50 ? 'PASSED' : 'FAILED';
  var rowNum   = sheet.getLastRow();

  sheet.appendRow([
    rowNum,
    data.timestamp || ts,
    data.name      || '',
    data.studentNo || 'N/A',
    data.section   || '',
    t1Pass ? '✅ PASS' : '❌ Not Passed',
    t1Pts,
    t2Pass ? '✅ PASS' : '❌ Not Passed',
    t2Pts,
    total,
    total + '%',
    grade,
    data.timeTaken || '—',
    data.autoSubmit ? '⏱ Auto (Time Up)' : '✅ Manual',
    data.task1Output || '(not attempted)',
    data.task2Output || '(not attempted)'
  ]);

  var lastRow = sheet.getLastRow();
  colorPracticalRow(sheet, lastRow, total, t1Pass, t2Pass);
}

// ─────────────────────────────────────────
//  THEORY HANDLER
// ─────────────────────────────────────────
function handleTheory(ss, data, ts) {
  var sheet = ss.getSheetByName(SHEET_THEORY);
  if (!sheet) { sheet = ss.insertSheet(SHEET_THEORY); setupTheorySheet(sheet); }

  var score  = data.score  || 0;
  var outOf  = data.outOf  || 20;
  var pct    = Math.round((score / outOf) * 100);
  var rowNum = sheet.getLastRow();

  sheet.appendRow([
    rowNum,
    data.timestamp || ts,
    data.name      || '',
    data.studentNo || 'N/A',
    data.section   || '',
    score, outOf, pct + '%',
    pct >= 50 ? 'PASSED' : 'FAILED',
    data.correct   || 0,
    data.wrong     || 0,
    data.timeTaken || '—',
    data.details   || ''
  ]);

  colorTheoryRow(sheet, sheet.getLastRow(), pct);
}

// ─────────────────────────────────────────
//  ACTIVITY LOG
// ─────────────────────────────────────────
function logActivity(ss, data, ts) {
  var sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(['Timestamp','Type','Name','Section','Score %','Status']);
    styleHeader(sheet, 1, '#37474F', '#FFFFFF');
    sheet.setFrozenRows(1);
    [180,160,220,100,90,180].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  }
  var status = data.type === 'DUPLICATE_BLOCKED' ? '🚫 Duplicate Blocked'
             : data.autoSubmit ? '⏱ Auto-Submitted' : '✅ Submitted';
  sheet.appendRow([ts, data.type||'—', data.name||'', data.section||'', data.percent||'—', status]);
}

// ─────────────────────────────────────────
//  STUDENT SUMMARY
// ─────────────────────────────────────────
function updateSummary(ss, data, ts) {
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SUMMARY);
    sheet.appendRow(['Full Name','Student No','Section','Theory Quiz %','Practical %','Combined Average','Last Updated']);
    styleHeader(sheet, 1, '#1A237E', '#FFFFFF');
    sheet.setFrozenRows(1);
    [220,130,110,130,130,150,200].forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  }
  var vals = sheet.getDataRange().getValues();
  var row  = -1;
  var name = (data.name || '').toLowerCase().trim();
  for (var i = 1; i < vals.length; i++) {
    if ((vals[i][0]||'').toString().toLowerCase().trim() === name) { row = i+1; break; }
  }
  if (row === -1) {
    sheet.appendRow([
      data.name||'', data.studentNo||'N/A', data.section||'',
      data.type==='PRACTICAL' ? '' : (data.percent||''),
      data.type==='PRACTICAL' ? (data.percent||'') : '',
      '', ts
    ]);
    row = sheet.getLastRow();
  } else {
    sheet.getRange(row, data.type==='PRACTICAL' ? 5 : 4).setValue(data.percent||'0%');
    sheet.getRange(row, 7).setValue(ts);
  }
  var t = parseFloat((sheet.getRange(row,4).getValue()||'').toString().replace('%',''))||0;
  var p = parseFloat((sheet.getRange(row,5).getValue()||'').toString().replace('%',''))||0;
  var hasBoth = sheet.getRange(row,4).getValue()!=='' && sheet.getRange(row,5).getValue()!=='';
  sheet.getRange(row,6).setValue(hasBoth ? ((t+p)/2).toFixed(1)+'%' : (sheet.getRange(row,4).getValue()||sheet.getRange(row,5).getValue()||''));
}

// ─────────────────────────────────────────
//  SHEET FORMATTING HELPERS
// ─────────────────────────────────────────
function setupPracticalSheet(sheet) {
  sheet.appendRow(PRACTICAL_HEADERS);
  styleHeader(sheet, 1, '#880E4F', '#FFFFFF');
  sheet.setFrozenRows(1);
  [50,190,220,130,100,120,90,120,90,100,100,90,110,150,280,280]
    .forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
  sheet.getRange(1,7).setBackground('#AD1457').setFontColor('#FFFFFF');
  sheet.getRange(1,9).setBackground('#AD1457').setFontColor('#FFFFFF');
  sheet.getRange(1,10).setBackground('#6A1B9A').setFontColor('#FFFFFF');
  sheet.getRange(1,11).setBackground('#6A1B9A').setFontColor('#FFFFFF');
  sheet.getRange(1,12).setBackground('#1B5E20').setFontColor('#FFFFFF');
}

function setupTheorySheet(sheet) {
  sheet.appendRow(THEORY_HEADERS);
  styleHeader(sheet, 1, '#1565C0', '#FFFFFF');
  sheet.setFrozenRows(1);
  [50,190,220,130,100,80,70,100,90,80,110,110,400]
    .forEach(function(w,i){ sheet.setColumnWidth(i+1,w); });
}

function styleHeader(sheet, rowNum, bg, fg) {
  var cols  = Math.max(sheet.getLastColumn(), 16);
  var range = sheet.getRange(rowNum, 1, 1, cols);
  range.setBackground(bg).setFontColor(fg).setFontWeight('bold')
       .setFontSize(10).setHorizontalAlignment('center');
}

function colorPracticalRow(sheet, row, pct, t1, t2) {
  var bg = pct>=100 ? '#E8F5E9' : pct>=50 ? '#FFF8E1' : '#FFEBEE';
  sheet.getRange(row,1,1,PRACTICAL_HEADERS.length).setBackground(bg).setFontSize(10);
  sheet.getRange(row,6).setBackground(t1?'#C8E6C9':'#FFCDD2').setFontWeight('bold');
  sheet.getRange(row,7).setBackground(t1?'#A5D6A7':'#EF9A9A').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(row,8).setBackground(t2?'#C8E6C9':'#FFCDD2').setFontWeight('bold');
  sheet.getRange(row,9).setBackground(t2?'#A5D6A7':'#EF9A9A').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(row,10).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(pct>=100?'#81C784':pct>=50?'#FFD54F':'#E57373')
    .setFontColor(pct>=50?'#1B5E20':'#B71C1C');
  sheet.getRange(row,12).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(pct>=50?'#2E7D32':'#C62828').setFontColor('#FFFFFF');
}

function colorTheoryRow(sheet, row, pct) {
  var bg = pct>=75 ? '#E8F5E9' : pct>=50 ? '#FFF8E1' : '#FFEBEE';
  sheet.getRange(row,1,1,THEORY_HEADERS.length).setBackground(bg).setFontSize(10);
  sheet.getRange(row,9).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground(pct>=50?'#2E7D32':'#C62828').setFontColor('#FFFFFF');
}

// ─────────────────────────────────────────
//  RUN ONCE — initial spreadsheet setup
// ─────────────────────────────────────────
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetName('CompProg 101 — Gradebook | DEBULGADO, S.M.G');

  var p = ss.getSheetByName(SHEET_PRACTICAL) || ss.insertSheet(SHEET_PRACTICAL);
  if (p.getLastRow()===0) setupPracticalSheet(p);

  var t = ss.getSheetByName(SHEET_THEORY) || ss.insertSheet(SHEET_THEORY);
  if (t.getLastRow()===0) setupTheorySheet(t);

  var s = ss.getSheetByName(SHEET_SUMMARY);
  if (!s) {
    s = ss.insertSheet(SHEET_SUMMARY);
    s.appendRow(['Full Name','Student No','Section','Theory Quiz %','Practical %','Combined Average','Last Updated']);
    styleHeader(s,1,'#1A237E','#FFFFFF'); s.setFrozenRows(1);
    [220,130,110,130,130,150,200].forEach(function(w,i){ s.setColumnWidth(i+1,w); });
  }

  var l = ss.getSheetByName(SHEET_LOG);
  if (!l) {
    l = ss.insertSheet(SHEET_LOG);
    l.appendRow(['Timestamp','Type','Name','Section','Score %','Status']);
    styleHeader(l,1,'#37474F','#FFFFFF'); l.setFrozenRows(1);
    [180,160,220,100,90,180].forEach(function(w,i){ l.setColumnWidth(i+1,w); });
  }

  try { var d=ss.getSheetByName('Sheet1'); if(d) ss.deleteSheet(d); } catch(e){}

  SpreadsheetApp.getUi().alert(
    '✅ Gradebook ready! 4 sheets created:\n\n' +
    '📋 practical_cfp_sheet — Task 1 & 2 results, 50pts each, 100% max\n' +
    '📝 Theory Quiz Results — 20-question quiz scores\n' +
    '👤 Student Summary — one row per student\n' +
    '📊 Activity Log — all submissions + blocked duplicates\n\n' +
    'Deploy → New Deployment → Web App → Anyone\n' +
    'Paste the URL into practical.html and quiz.html'
  );
}
