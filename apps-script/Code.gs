/**
 * Half Marathon Plan: Google Sheet + sync for https://pengkistan.github.io/marathonplan/
 *
 * setup()  builds (or rebuilds) the sheets from plan.json and keeps anything already logged.
 * doGet()  returns every logged run, test times and max heart rate as JSON.
 * doPost() saves runs and settings sent by the app, then returns the same JSON.
 * onEdit() stamps rows you edit by hand so the app picks up your changes.
 */
const PLAN_URL = 'https://pengkistan.github.io/marathonplan/plan.json';
const APP_URL = 'https://pengkistan.github.io/marathonplan/';
const DOC_URL = 'https://claude.ai/artifact/BErnvi4zztgJg9Htm676nq';
const SH = { start: 'Start here', plan: 'Plan', week: 'Weekly', pace: 'Paces', gym: 'Gym' };
// Plan sheet columns (1-based)
const C = { date: 1, week: 2, phase: 3, day: 4, session: 5, workout: 6, planned: 7, vdot: 8, target: 9,
  done: 10, km: 11, time: 12, hr: 13, feel: 14, pace: 15, notes: 16, key: 17, id: 18, ts: 19 };
const PHASE = { base: 'Base', build: 'Build', specific: 'Race-specific', taper: 'Taper & race' };
// Daniels' Running Formula, tables 5.1 and 5.2 (VDOT 30-40)
const PACES = [
  [30, '30:40', '63:46', '2:21:04', '6:24', '5:55', '2:22', 67, 101, '7:03', '7:27–8:14'],
  [31, '29:51', '62:03', '2:17:21', '6:14', '5:45', '2:18', 65, 98, '6:52', '7:16–8:02'],
  [32, '29:05', '60:26', '2:13:49', '6:05', '5:35', '2:14', 63, 95, '6:40', '7:05–7:52'],
  [33, '28:21', '58:54', '2:10:27', '5:56', '5:28', '2:11', 61, 92, '6:30', '6:55–7:41'],
  [34, '27:39', '57:26', '2:07:16', '5:48', '5:20', '2:08', 60, 90, '6:20', '6:45–7:31'],
  [35, '27:00', '56:03', '2:04:13', '5:40', '5:13', '2:05', 58, 87, '6:10', '6:36–7:21'],
  [36, '26:22', '54:44', '2:01:19', '5:33', '5:05', '2:02', 57, 85, '6:01', '6:27–7:11'],
  [37, '25:46', '53:29', '1:58:34', '5:26', '5:00', '1:59', 55, 83, '5:53', '6:19–7:02'],
  [38, '25:12', '52:17', '1:55:55', '5:19', '4:54', '1:56', 54, 81, '5:45', '6:11–6:54'],
  [39, '24:39', '51:09', '1:53:24', '5:12', '4:48', '1:54', 53, 80, '5:37', '6:03–6:46'],
  [40, '24:08', '50:03', '1:50:59', '5:06', '4:42', '1:52', 52, 78, '5:29', '5:56–6:38'],
];
const NAVY = '#16213A', YELLOW = '#FFF2CC', MUTED = '#5B667E', FONT = 'Arial';

/* ---------------- helpers ---------------- */
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}
function secs_(s) {
  if (s === null || s === undefined || s === '') return null;
  const p = String(s).trim().split(':').map(Number);
  if (p.some(isNaN)) return null;
  return p.length === 3 ? p[0] * 3600 + p[1] * 60 + p[2] : p.length === 2 ? p[0] * 60 + p[1] : p[0] * 60;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(ss, name, index) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name, index);
  return sh;
}
function head_(range) {
  range.setFontFamily(FONT).setFontWeight('bold').setFontColor('#FFFFFF').setBackground(NAVY)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

/* ---------------- web app ---------------- */
function doGet() {
  return out_(Object.assign({ ok: true }, readAll_(book_())));
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Request was not valid JSON' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = book_();
    if (body.log) writeLogs_(ss, body.log);
    if (body.settings) writeSettings_(ss, body.settings);
    SpreadsheetApp.flush();
    return out_(Object.assign({ ok: true }, readAll_(ss)));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function readAll_(ss) {
  const sh = ss.getSheetByName(SH.plan), st = ss.getSheetByName(SH.start);
  const n = sh.getLastRow() - 1;
  const raw = sh.getRange(2, 1, n, C.ts).getValues();
  const disp = sh.getRange(2, 1, n, C.ts).getDisplayValues();
  const log = {};
  for (let i = 0; i < n; i++) {
    const r = raw[i], d = disp[i], id = r[C.id - 1];
    if (!id) continue;
    const done = String(r[C.done - 1]).trim().toLowerCase() === 'yes';
    const km = r[C.km - 1] === '' ? null : Number(r[C.km - 1]);
    const time = d[C.time - 1] || '';
    const hr = r[C.hr - 1] === '' ? null : Number(r[C.hr - 1]);
    const feel = r[C.feel - 1] === '' ? null : Number(r[C.feel - 1]);
    const notes = String(r[C.notes - 1] || '');
    const ts = Number(r[C.ts - 1]) || 0;
    if (done || km !== null || time || hr !== null || feel !== null || notes || ts) {
      log[id] = { done: done, km: km, time: time, hr: hr, feel: feel, notes: notes, ts: ts };
    }
  }
  const s = st.getRange('B6:C9').getValues();
  const t = (m, sec) => (m === '' || m === null) ? '' : `${Number(m)}:${String(Number(sec) || 0).padStart(2, '0')}`;
  return { log: log, tests: { t5k: t(s[2][0], s[2][1]), t10k: t(s[3][0], s[3][1]) }, hrMax: Number(s[0][0]) || 194, sheetUrl: ss.getUrl() };
}

function writeLogs_(ss, log) {
  const sh = ss.getSheetByName(SH.plan);
  const n = sh.getLastRow() - 1;
  const ids = sh.getRange(2, C.id, n, 1).getValues().map(r => r[0]);
  const tss = sh.getRange(2, C.ts, n, 1).getValues().map(r => Number(r[0]) || 0);
  Object.keys(log).forEach(id => {
    const i = ids.indexOf(id);
    if (i < 0) return;
    const L = log[id] || {};
    const ts = Number(L.ts) || Date.now();
    if (ts < tss[i]) return; // the sheet has a newer edit
    const row = i + 2, sec = secs_(L.time);
    sh.getRange(row, C.done, 1, 5).setValues([[L.done ? 'Yes' : '', L.km == null ? '' : Number(L.km),
      sec ? sec / 86400 : '', L.hr == null ? '' : Number(L.hr), L.feel == null ? '' : Number(L.feel)]]);
    sh.getRange(row, C.notes).setValue(L.notes || '');
    sh.getRange(row, C.ts).setValue(ts);
  });
}

function writeSettings_(ss, s) {
  const st = ss.getSheetByName(SH.start);
  if (s.hrMax) st.getRange('B6').setValue(Number(s.hrMax));
  const put = (key, row) => {
    if (!(key in s)) return;
    const sec = secs_(s[key]);
    st.getRange(row, 2, 1, 2).setValues(sec ? [[Math.floor(sec / 60), sec % 60]] : [['', '']]);
  };
  put('t5k', 8); put('t10k', 9);
}

/* ---------------- edits made in the sheet ---------------- */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SH.plan) return;
  const r = e.range.getRow(), c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  if (r < 2 || c2 < C.done || c1 > C.notes) return;
  const rows = e.range.getNumRows(), now = Date.now();
  sh.getRange(r, C.ts, rows, 1).setValues(Array.from({ length: rows }, () => [now]));
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Half marathon')
    .addItem('Rebuild plan sheets (keeps your logs)', 'setup')
    .addItem('Open the web app', 'openApp_')
    .addToUi();
}
function openApp_() {
  const html = HtmlService.createHtmlOutput(`<p style="font-family:Arial"><a href="${APP_URL}" target="_blank">Open the training app</a></p>`).setWidth(260).setHeight(60);
  SpreadsheetApp.getUi().showModalDialog(html, 'Half marathon');
}

/* ---------------- build the sheets ---------------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  // Formulas below use English names and commas, so the sheet uses the US number format.
  ss.setSpreadsheetLocale('en_US');
  ss.setSpreadsheetTimeZone('Asia/Makassar');
  SpreadsheetApp.flush();
  const plan = JSON.parse(UrlFetchApp.fetch(PLAN_URL).getContentText());

  // keep anything already logged
  let keepLog = {}, keepInputs = null;
  if (ss.getSheetByName(SH.plan) && ss.getSheetByName(SH.plan).getLastRow() > 1) keepLog = readAll_(ss).log;
  if (ss.getSheetByName(SH.start)) keepInputs = ss.getSheetByName(SH.start).getRange('B6:C9').getValues();

  ss.rename('Half Marathon Plan 2027');
  const first = ss.getSheets()[0];
  const st = sheet_(ss, SH.start, 0), pl = sheet_(ss, SH.plan, 1), wk = sheet_(ss, SH.week, 2),
    ps = sheet_(ss, SH.pace, 3), gy = sheet_(ss, SH.gym, 4);
  [st, pl, wk, ps, gy].forEach(sh => { sh.clear(); sh.clearConditionalFormatRules(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); });
  if ([st, pl, wk, ps, gy].indexOf(first) < 0 && ss.getSheets().length > 5) ss.deleteSheet(first);

  buildPaces_(ps);
  buildStart_(st, keepInputs);
  const last = buildPlan_(pl, plan);
  buildWeekly_(wk, plan, last);
  buildGym_(gy);
  if (Object.keys(keepLog).length) writeLogs_(ss, keepLog);
  ss.setActiveSheet(st);
}

function buildPaces_(ps) {
  ps.getRange('A1').setValue('Training paces and race times by VDOT').setFontWeight('bold').setFontSize(14).setFontColor(NAVY);
  ps.getRange('A2').setValue("Source: Daniels' Running Formula, 4th ed., tables 5.1 and 5.2. HM pace = half-marathon time / 21.0975 km.")
    .setFontStyle('italic').setFontSize(9).setFontColor(MUTED);
  const heads = ['VDOT', '5K', '10K', 'Half', 'HM pace', 'T pace', 'I pace', 'R 200 m', 'R 300 m', 'M pace', 'Easy (Daniels)'];
  ps.getRange(4, 1, 1, heads.length).setValues([heads]); head_(ps.getRange(4, 1, 1, heads.length));
  const rows = PACES.map(p => [p[0], secs_(p[1]) / 86400, secs_(p[2]) / 86400, secs_(p[3]) / 86400, '',
    `${p[4]}/km`, `${p[5]}/km (${p[6]} per 400 m)`, `${p[7]} s per 200 m`, `${p[8]} s per 300 m`, `${p[9]}/km`, `${p[10]}/km`]);
  ps.getRange(5, 1, rows.length, 11).setValues(rows);
  ps.getRange(5, 5, rows.length, 1).setFormulas(rows.map((_, i) => {
    const r = i + 5;
    return [`=INT(ROUND(D${r}*86400/21.0975,0)/60)&":"&TEXT(MOD(ROUND(D${r}*86400/21.0975,0),60),"00")&"/km"`];
  }));
  ps.getRange(5, 2, rows.length, 2).setNumberFormat('[m]:ss');
  ps.getRange(5, 4, rows.length, 1).setNumberFormat('[h]:mm:ss');
  ps.getRange(5, 1, rows.length, 11).setFontFamily(FONT).setHorizontalAlignment('center');
  [60, 70, 70, 80, 85, 80, 190, 120, 120, 80, 120].forEach((w, i) => ps.setColumnWidth(i + 1, w));
  ps.setFrozenRows(4);
}

function buildStart_(st, keep) {
  st.setColumnWidth(1, 300); st.setColumnWidth(2, 100); st.setColumnWidth(3, 90); st.setColumnWidth(4, 90); st.setColumnWidth(5, 380);
  st.getRange('A1').setValue('Half Marathon Plan, race on Sat Feb 6, 2027').setFontSize(16).setFontWeight('bold').setFontColor(NAVY);
  st.getRange('A2').setValue('15 weeks from Mon Oct 26, 2026. The web app saves every run you log into the Plan sheet.').setFontColor(MUTED);
  st.getRange('A3').setValue('Days to race').setFontWeight('bold');
  st.getRange('B3').setFormula('=MAX(0,DATE(2027,2,6)-TODAY())').setFontWeight('bold').setFontSize(14);
  st.getRange('A5').setValue('Your inputs (yellow cells; the app fills these too)').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  st.getRange('A6:A9').setValues([['Max heart rate (bpm)'], ['Starting VDOT (Base phase)'], ['5K test, Sat Nov 21: minutes / seconds'], ['10K test, Sat Dec 19: minutes / seconds']]);
  st.getRange('B6:C9').setValues(keep || [[194, ''], [32, ''], ['', ''], ['', '']]);
  st.getRange('B6:B9').setBackground(YELLOW).setFontColor('#0000FF').setHorizontalAlignment('center');
  st.getRange('C8:C9').setBackground(YELLOW).setFontColor('#0000FF').setHorizontalAlignment('center');
  st.getRange('E6:E9').setValues([['Highest seen in your Strava data (10K, Aug 7, 2026).'], ['From your 5K 28:26 (Feb 4) and 10K 59:54 (Aug 7).'],
    ['Example: 27 and 50 for 27:50.'], ['Example: 58 and 30 for 58:30; 61 and 5 for 1:01:05.']]).setFontSize(9).setFontColor(MUTED);
  st.getRange('A11').setValue('What your tests say').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  st.getRange('A12:A17').setValues([['VDOT from 5K test'], ['VDOT from 10K test'], ['Build-phase VDOT (weeks 5–9)'], ['Race-phase VDOT (weeks 10–15)'], ['Predicted half marathon'], ['Half-marathon pace']]);
  st.getRange('B12:B17').setFormulas([
    ['=IF(N(B8)=0,"",MAX(30,29+SUMPRODUCT(--(Paces!$B$5:$B$15>=TIME(0,B8,N(C8))))))'],
    ['=IF(N(B9)=0,"",MAX(30,29+SUMPRODUCT(--(Paces!$C$5:$C$15>=TIME(0,B9,N(C9))))))'],
    ['=IF(B12="",B7,B12)'], ['=IF(B13="",B14,B13)'],
    ['=INDEX(Paces!$D$5:$D$15,MATCH(B15,Paces!$A$5:$A$15,0))'], ['=INDEX(Paces!$E$5:$E$15,MATCH(B15,Paces!$A$5:$A$15,0))']]).setFontWeight('bold').setHorizontalAlignment('center');
  st.getRange('B16').setNumberFormat('[h]:mm:ss');
  st.getRange('E12').setValue('Rounded down to the VDOT row your time reaches (Daniels, table 5.1).').setFontSize(9).setFontColor(MUTED);
  st.getRange('E16').setValue('Plan target: 2:07:16 (VDOT 34). Stretch: 2:04:13 (VDOT 35).').setFontSize(9).setFontColor(MUTED);
  st.getRange('A19').setValue('How it works').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  st.getRange('A20:A24').setValues([
    ['1. Log runs in the web app; each one lands in the Plan sheet within a few seconds.'],
    ['2. You can also type in the yellow Plan columns here; the app picks the change up next time it opens.'],
    ['3. Type time as h:mm:ss, for example 0:41:20 (41:20 would mean 41 hours).'],
    ['4. Weekly totals planned and logged km for each week.'],
    ['5. Move up a VDOT only after a test, and not more often than every third week (Daniels).']]);
  st.getRange('A26').setValue('Links').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  st.getRange('A27:B28').setValues([['Web app', APP_URL], ['Full plan with sources', DOC_URL]]);
  st.getRange(1, 1, 30, 5).setFontFamily(FONT);
}

function buildPlan_(pl, plan) {
  const heads = ['Date', 'Week', 'Phase', 'Day', 'Session', 'Workout', 'Planned km', 'VDOT', 'Target', 'Done', 'Actual km', 'Time', 'Avg HR', 'Feel (1–5)', 'Pace /km', 'Notes', 'Key', 'ID', 'Updated'];
  pl.getRange(1, 1, 1, heads.length).setValues([heads]); head_(pl.getRange(1, 1, 1, heads.length));
  const S = plan.sessions, n = S.length, last = n + 1;
  const keyOf = s => s.type === 'R' ? (s.detail.indexOf('300 m') >= 0 ? 'R 300 m' : 'R 200 m')
    : ({ T: 'T pace', I: 'I pace', M: 'M pace', HM: 'HM pace', E: 'E', REC: 'REC', L: 'L' }[s.pace] || '');
  const vals = S.map(s => {
    return [s.date, s.week, PHASE[s.phase], s.day, s.title, s.detail, s.type === 'REST' ? '' : s.km];
  });
  pl.getRange(2, 1, n, 7).setValues(vals);
  pl.getRange(2, C.key, n, 3).setValues(S.map(s => [keyOf(s), s.id, '']));
  const hr = "'Start here'!$B$6";
  pl.getRange(2, C.vdot, n, 2).setFormulas(S.map((_, i) => {
    const r = i + 2;
    return [`=IF(C${r}="Base",'Start here'!$B$7,IF(C${r}="Build",'Start here'!$B$14,'Start here'!$B$15))`,
      `=IF(Q${r}="","",IF(Q${r}="E","HR under "&ROUND(${hr}*0.79,0),IF(Q${r}="REC","HR under "&ROUND(${hr}*0.76,0),IF(Q${r}="L","HR "&ROUND(${hr}*0.73,0)&" to "&ROUND(${hr}*0.84,0),INDEX(Paces!$B$5:$K$15,MATCH(H${r},Paces!$A$5:$A$15,0),MATCH(Q${r},Paces!$B$4:$K$4,0))))))`];
  }));
  pl.getRange(2, C.pace, n, 1).setFormulas(S.map((_, i) => [`=IF(AND(N(K${i + 2})>0,N(L${i + 2})>0),L${i + 2}/K${i + 2},"")`]));
  pl.getRange(2, 1, n, 1).setNumberFormat('ddd d mmm');
  pl.getRange(2, C.planned, n, 1).setNumberFormat('0.0');
  pl.getRange(2, C.km, n, 1).setNumberFormat('0.0');
  pl.getRange(2, C.time, n, 1).setNumberFormat('[h]:mm:ss');
  pl.getRange(2, C.pace, n, 1).setNumberFormat('[m]:ss');
  pl.getRange(2, C.ts, n, 1).setNumberFormat('0');
  pl.getRange(2, 1, n, C.ts).setFontFamily(FONT).setFontSize(10).setVerticalAlignment('top');
  pl.getRange(2, C.workout, n, 1).setWrap(true);
  pl.getRange(2, C.notes, n, 1).setWrap(true);
  pl.getRange(2, C.done, n, 5).setBackground(YELLOW);
  pl.getRange(2, C.notes, n, 1).setBackground(YELLOW);
  [90, 50, 100, 45, 160, 380, 75, 50, 190, 55, 75, 75, 60, 65, 65, 220, 70, 70, 110].forEach((w, i) => pl.setColumnWidth(i + 1, w));
  pl.hideColumns(C.key, 3);
  pl.setFrozenRows(1); pl.setFrozenColumns(1);
  pl.getRange(2, C.done, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build());
  pl.getRange(2, C.feel, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 5).setAllowInvalid(false).setHelpText('1 very easy to 5 very hard').build());
  pl.getRange(2, C.hr, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(60, 230).setAllowInvalid(false).build());
  const rng = pl.getRange(2, 1, n, C.target);
  pl.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2="Yes"').setBackground('#D9EEEC').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2=TODAY()').setBackground('#FCE4D6').setRanges([rng]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$E2="Rest"').setFontColor('#9AA3B5').setRanges([rng]).build(),
  ]);
  S.forEach((s, i) => { if (s.type === 'TEST' || s.type === 'RACE' || s.type === 'REH') pl.getRange(i + 2, C.session, 1, 2).setFontWeight('bold'); });
  return last;
}

function buildWeekly_(wk, plan, last) {
  wk.getRange('A1').setValue('Planned and logged distance by week').setFontSize(14).setFontWeight('bold').setFontColor(NAVY);
  wk.getRange('A2').setValue('Updates from the Done and Actual km columns on the Plan sheet. Week 15 includes the race (21.1 km).').setFontSize(9).setFontColor(MUTED);
  const heads = ['Week', 'Phase', 'From', 'To', 'Planned km', 'Runs planned', 'Runs done', 'Logged km', 'Logged / planned'];
  wk.getRange(4, 1, 1, heads.length).setValues([heads]); head_(wk.getRange(4, 1, 1, heads.length));
  const W = plan.weeks;
  wk.getRange(5, 1, W.length, 4).setValues(W.map(w => [w.week, PHASE[w.phase], w.start, w.end]));
  wk.getRange(5, 5, W.length, 5).setFormulas(W.map((_, i) => {
    const r = i + 5;
    return [`=SUMIFS(Plan!$G$2:$G$${last},Plan!$B$2:$B$${last},A${r})`,
      `=COUNTIFS(Plan!$B$2:$B$${last},A${r},Plan!$E$2:$E$${last},"<>Rest")`,
      `=COUNTIFS(Plan!$B$2:$B$${last},A${r},Plan!$J$2:$J$${last},"Yes")`,
      `=SUMIFS(Plan!$K$2:$K$${last},Plan!$B$2:$B$${last},A${r},Plan!$J$2:$J$${last},"Yes")`,
      `=IF(E${r}>0,H${r}/E${r},0)`];
  }));
  const tr = 5 + W.length;
  wk.getRange(tr, 1).setValue('Total').setFontWeight('bold');
  wk.getRange(tr, 5, 1, 5).setFormulas([[`=SUM(E5:E${tr - 1})`, `=SUM(F5:F${tr - 1})`, `=SUM(G5:G${tr - 1})`, `=SUM(H5:H${tr - 1})`, `=IF(E${tr}>0,H${tr}/E${tr},0)`]]).setFontWeight('bold');
  wk.getRange(5, 3, W.length, 2).setNumberFormat('d mmm');
  wk.getRange(5, 5, W.length + 1, 1).setNumberFormat('0.0');
  wk.getRange(5, 8, W.length + 1, 1).setNumberFormat('0.0');
  wk.getRange(5, 9, W.length + 1, 1).setNumberFormat('0%');
  wk.getRange(5, 1, W.length + 1, 9).setFontFamily(FONT).setHorizontalAlignment('center');
  wk.getCharts().forEach(c => wk.removeChart(c));
  const chart = wk.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(wk.getRange(4, 1, W.length + 1, 1)).addRange(wk.getRange(4, 5, W.length + 1, 1)).addRange(wk.getRange(4, 8, W.length + 1, 1))
    .setNumHeaders(1).setPosition(4, 11, 0, 0)
    .setOption('title', 'Planned vs logged km per week').setOption('colors', ['#C9D0DC', '#237A7F'])
    .setOption('legend', { position: 'bottom' }).setOption('width', 620).setOption('height', 320).build();
  wk.insertChart(chart);
}

function buildGym_(gy) {
  gy.setColumnWidth(1, 280); gy.setColumnWidth(2, 170); gy.setColumnWidth(3, 420);
  let r = 1;
  const title = t => { gy.getRange(r, 1).setValue(t).setFontWeight('bold').setFontSize(13).setFontColor(NAVY); r++; };
  const table = rows => { gy.getRange(r, 1, rows.length, 3).setValues(rows).setWrap(true).setVerticalAlignment('top'); head_(gy.getRange(r, 1, 1, 3)); r += rows.length + 1; };
  title('Gym A: heavy (Sunday, after the recovery run, about 40 min)');
  table([['Exercise', 'Sets × reps', 'Notes'], ['Squat (back or goblet) or trap-bar deadlift', '3 × 5', 'Heavy: you could do only about 2 more'],
    ['Romanian deadlift', '3 × 6', ''], ['Bulgarian split squat', '3 × 6 per leg', ''], ['Standing calf raise', '3 × 8', 'Heavy'],
    ['Bent-knee calf raise', '3 × 10', 'Works the soleus'], ['Core: plank, side plank, dead bug', '2 rounds', '']]);
  title('Gym B: power (Wednesday, after the easy run, about 25 min)');
  table([['Exercise', 'Sets × reps', 'Notes'], ['Pogo hops', '3 × 20 contacts', 'Quick, stiff ankles'], ['Box jumps', '3 × 5', "Step down, don't jump down"],
    ['Bounding', '3 × 20 m', 'From week 5'], ['Single-leg Romanian deadlift', '3 × 8 per leg', ''],
    ['Isometric hold (optional)', '6 × 5 s', 'Push hard against a fixed bar, split-squat or calf-raise position'], ['Core: plank, side plank, dead bug', '2 rounds', '']]);
  title('How the gym work progresses');
  table([['Weeks', 'What to do', ''], ['1–2', 'Learn the movements', 'Moderate weights, 3 × 8 in Gym A'], ['3–9', 'Go heavy', '3 × 4–6; skip the sessions marked in test weeks 4 and 8'],
    ['10–13', 'Maintain', 'Same weights, 2 sets; half the jumps in Gym B'], ['14–15', 'Taper', 'Core only in week 14, nothing in race week']]);
  title('Hill sprints and strides');
  table([['What', 'How', 'Progression'], ['Hill sprints (Hudson, Snow)', '6–8% hill, 8–10 s all-out, walk down, rest at least 3 min', '2, 4, 6 × 8 s (weeks 1–3); 8 × 8 s (week 5); 8 × 10 s (weeks 6–7); 6 × 10 s to Jan 20'],
    ['Strides (Daniels)', '15–20 s light and quick at about R pace; 45–60 s easy between', 'Never a sprint']]);
  gy.getRange(1, 1, r, 3).setFontFamily(FONT);
}
