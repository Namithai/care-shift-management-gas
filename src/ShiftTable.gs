/**
 * 老健シフト管理アプリ — シフト表画面（第3弾）
 * shift シート: A=年月 / B=スタッフNo / C〜AG=1日〜31日
 */

/** 動作確認用。エディタから実行して実行ログを見る */
function testShiftIssues() {
  var r = getShiftIssues('2026-09', { maxRun: 6, needLeader: true });
  Logger.log('人が足りない枠: ' + r.lackTotal + '件 / 区分' + r.lack.length + '種類');
  for (var i = 0; i < r.lack.length; i++) Logger.log('  ' + r.lack[i].label + ' : ' + r.lack[i].days.join(','));
  Logger.log('夜勤リーダーがいない日: ' + r.noLead.join(','));
  Logger.log('公休不足: ' + r.offShort.map(function (x) { return x.name + '(' + x.days + '/' + x.quota + ')'; }).join(' , '));
  Logger.log('連勤超過: ' + r.runOver.map(function (x) { return x.name + '(' + x.run + '連勤 ' + x.start + '-' + x.end + ')'; }).join(' , '));
  Logger.log('空きマス: ' + r.blank.map(function (x) { return x.name + '(' + x.days + ')'; }).join(' , '));
  Logger.log('hasIssue: ' + r.hasIssue);
}



var SHIFT_SHEET = 'shift';
var SHIFT_FIXED_COLS = 2;
var SHIFT_MAX_DAYS = 31;

var FLOOR_LABEL = floorLabelMap_();   // floor シートから作る
var FLOOR_ORDER = floorCodes_();

// 現行の勤務表と同じ並び（表示グループ）
var GROUP_ORDER = floorGroups_();   // floor シートの「グループ」から作る

/** ------------------------------------------------------------------
 *  初回セットアップ（1回だけ実行する）
 *  ------------------------------------------------------------------ */
function setupShiftSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHIFT_SHEET);
  if (!sh) sh = ss.insertSheet(SHIFT_SHEET);

  var head = ['年月', 'スタッフNo'];
  for (var d = 1; d <= SHIFT_MAX_DAYS; d++) head.push(String(d));

  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(SHIFT_FIXED_COLS);
  sh.getRange(1, 1, sh.getMaxRows(), head.length)
    .setHorizontalAlignment('center')
    .setNumberFormat('@');
  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(2, 90);
  for (var c = 3; c <= head.length; c++) sh.setColumnWidth(c, 34);

  SpreadsheetApp.getUi().alert('shift シートを用意しました。');
}

/** 動作確認用のサンプルデータ（2026-09、ダミー3名） */
function setupShiftSampleData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHIFT_SHEET);
  if (!sh) { SpreadsheetApp.getUi().alert('先に setupShiftSheet を実行してください。'); return; }

  var ym = '2026-09';
  var days = daysInMonth_(ym);

  var plan = {
    'S001': ['A3','A3','夜','明','休','A3','A3','A3','夜','明','休','希','A3','A3','夜','明','休','A3','A3','A3','夜','明','休','希','A3','A3','夜','明','休','希'],
    'S002': ['F9','F9','休','F9','F9','休','休','F9','F9','休','F9','F9','休','休','F9','F9','休','F9','F9','休','休','F9','F9','休','F9','F9','休','希','希','休'],
    'S003': ['F','F','休','F','F','休','休','F','F','休','F','F','休','休','F','F','休','F','F','休','休','F','F','休','F','F','休','休','F','F']
  };

  var rows = [];
  Object.keys(plan).forEach(function (no) {
    var row = [ym, no];
    for (var d = 0; d < SHIFT_MAX_DAYS; d++) row.push(d < days ? (plan[no][d] || '') : '');
    rows.push(row);
  });

  var last = sh.getLastRow();
  if (last > 1) {
    var cur = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = cur.length - 1; i >= 0; i--) {
      if (String(cur[i][0]) === ym) sh.deleteRow(i + 2);
    }
  }
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, SHIFT_FIXED_COLS + SHIFT_MAX_DAYS).setValues(rows);
  SpreadsheetApp.getUi().alert('サンプルデータ（' + ym + '・3名）を入れました。');
}

/** ------------------------------------------------------------------
 *  画面を開く
 *  ------------------------------------------------------------------ */
function openShiftTable() {
  var t = HtmlService.createTemplateFromFile('shift');
  t.nav = '';
  var html = t.evaluate()
    .setWidth(2000)
    .setHeight(1400);
  SpreadsheetApp.getUi().showModalDialog(html, 'シフト表');
}

/** ------------------------------------------------------------------
 *  データ取得
 *  ------------------------------------------------------------------ */

function getCurrentYearMonth() {
  var now = new Date();
  return Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM');
}

function getShiftMonth(ym) {
  if (!ym) ym = getCurrentYearMonth();
  var days = daysInMonth_(ym);
  var staffList = readStaffForShift_();
  var shiftMap = readShiftRows_(ym);

  var rows = staffList.map(function (s) {
    var cells = shiftMap[s.no] || [];
    var arr = [];
    for (var d = 0; d < days; d++) arr.push(String(cells[d] || ''));
    return {
      no: s.no,
      name: s.name,
      jobType: s.jobType,
      floor: s.floor,
      subFloors: s.subFloors,
      group: s.group,
      countable: s.countable,
      employment: s.employment,
      cells: arr,
      total: countStaffMonth_(arr)
    };
  });

  return {
    ym: ym,
    label: ym.replace('-', '年') + '月',
    days: days,
    weekdays: weekdayLabels_(ym, days),
    requiredHoliday: requiredHolidayFor_(ym, days),
    groups: GROUP_ORDER.slice(),
    rows: rows,
    symbols: getSymbolList_(),
    staffing: getStaffing_(),
    daily: countDaily_(rows, days)
  };
}

function readStaffForShift_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('staff');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 20).getValues();
  var out = [];
  v.forEach(function (r) {
    var no = String(r[0] || '').trim();
    if (!no) return;
    if (String(r[17] || '在籍') === '退職') return;
    out.push({
      no: no,
      name: String(r[1] || ''),
      floor: String(r[3] || ''),
      subFloors: String(r[4] || ''),
      countable: !(r[16] === true || String(r[16]) === 'TRUE'),
      jobType: String(r[18] || ''),
      group: String(r[19] || ''),
      employment: String(r[5] || '')
    });
  });
  return out;
}

function readShiftRows_(ym) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHIFT_SHEET);
  var map = {};
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  var v = sh.getRange(2, 1, last - 1, SHIFT_FIXED_COLS + SHIFT_MAX_DAYS).getValues();
  v.forEach(function (r) {
    if (String(r[0]) !== ym) return;
    map[String(r[1]).trim()] = r.slice(SHIFT_FIXED_COLS);
  });
  return map;
}

/** ------------------------------------------------------------------
 *  記号の判定と集計
 *  ------------------------------------------------------------------ */

function classifySymbol(sym) {
  var s = String(sym || '').trim();
  if (!s) return '';
  var sp = s.indexOf(' ');
  if (sp > 0) s = s.slice(0, sp); // 「A 5」のように後ろにフロアが付く場合は記号だけ見る
  var m = getSymbolMap_()[s];
  if (m && KIND_KEY[m.kind]) return KIND_KEY[m.kind];
  if (s === '休') return 'off';
  if (s === '有') return 'paid';
  if (s === '希') return 'wish';
  if (s === '明') return 'after';
  if (s === '夜' || s === '新夜') return 'night';
  var head = s.charAt(0);
  if (head === 'A') return 'early';
  if (head === 'B') return 'late';
  if (head === 'C') return 'day';
  if (head === 'F') return 'free';
  return 'other';
}

function countStaffMonth_(cells) {
  var t = { work: 0, night: 0, off: 0, paid: 0 };
  cells.forEach(function (c) {
    var k = classifySymbol(c);
    if (k === 'off' || k === 'wish') t.off++;
    else if (k === 'paid') t.paid++;
    else if (k === 'night') { t.night++; t.work++; }
    else if (k === 'early' || k === 'late' || k === 'day' || k === 'free' || k === 'other' || k === 'after') t.work++; // 明けも出勤日数に数える（画面側 recount と同じ）
  });
  return t;
}

function countDaily_(rows, days) {
  var keys = ['early', 'late', 'day', 'free', 'night', 'off'];
  var out = {};
  keys.forEach(function (k) {
    out[k] = [];
    for (var i = 0; i < days; i++) out[k].push(0);
  });
  rows.forEach(function (r) {
    if (!r.countable) return;
    for (var d = 0; d < days; d++) {
      var k = classifySymbol(r.cells[d]);
      if (k === 'wish') k = 'off';
      if (out[k]) out[k][d]++;
    }
  });
  return out;
}

/** ------------------------------------------------------------------
 *  日付まわり
 *  ------------------------------------------------------------------ */
function daysInMonth_(ym) {
  var p = String(ym).split('-');
  return new Date(Number(p[0]), Number(p[1]), 0).getDate();
}

function weekdayLabels_(ym, days) {
  var p = String(ym).split('-');
  var names = ['日', '月', '火', '水', '木', '金', '土'];
  var out = [];
  for (var d = 1; d <= days; d++) {
    var w = new Date(Number(p[0]), Number(p[1]) - 1, d).getDay();
    out.push({ label: names[w], dow: w });
  }
  return out;
}


/** ------------------------------------------------------------------
 *  記号マスタ・必要人数マスタ・保存
 *  ------------------------------------------------------------------ */

var KIND_KEY = { '早番': 'early', '遅番': 'late', '日勤': 'day', 'F勤': 'free', '夜勤': 'night', '明け': 'after', '公休': 'off', '有休': 'paid', '希望休': 'wish', '看護': 'other' };

var SYMBOL_LIST_CACHE = null;
var SYMBOL_MAP_CACHE = null;

/** symbol シートを読む（1回の実行につき1度だけ） */
function getSymbolList_() {
  if (SYMBOL_LIST_CACHE) return SYMBOL_LIST_CACHE;
  var out = [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('symbol');
  if (sh) {
    var last = sh.getLastRow();
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < v.length; i++) {
        var sym = String(v[i][0] || '').trim();
        if (!sym) continue;
        out.push({ sym: sym, kind: String(v[i][1] || '').trim(), floor: String(v[i][2] || '').trim(), note: String(v[i][3] || '').trim() });
      }
    }
  }
  SYMBOL_LIST_CACHE = out;
  return out;
}

function getSymbolMap_() {
  if (SYMBOL_MAP_CACHE) return SYMBOL_MAP_CACHE;
  var map = {};
  getSymbolList_().forEach(function (o) { map[o.sym] = o; });
  SYMBOL_MAP_CACHE = map;
  return map;
}

/** staffing シート（必要人数）を読む */
function getStaffing_() {
  var out = [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('staffing');
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    var k = String(v[i][0] || '').trim();
    var f = String(v[i][1] || '').trim();
    if (!k || !f) continue;
    var mn = v[i][2], wt = v[i][3];
    // 1つのマスに「1<タブ>2」のように2つ入っている場合は分ける
    if (String(wt) === '') {
      var TAB = String.fromCharCode(9);
      var parts = String(mn).split(TAB).join(' ').split('　').join(' ').split(',').join(' ').split('、').join(' ').split('/').join(' ').split(' ');
      var nums = [];
      for (var pi = 0; pi < parts.length; pi++) if (parts[pi] !== '') nums.push(parts[pi]);
      if (nums.length >= 2) { mn = nums[0]; wt = nums[1]; }
    }
    var mnum = (String(mn) === '' ? null : Number(mn));
    var wnum = (String(wt) === '' ? null : Number(wt));
    if (mnum !== null && isNaN(mnum)) mnum = null;
    if (wnum !== null && isNaN(wnum)) wnum = null;
    if (mnum === null && wnum === null) continue;
    out.push({ kind: k, floor: f, min: mnum, want: wnum });
  }
  return out;
}

/**
 * シフト表のマスをまとめて保存する。
 * changes = [{ no: 'S001', day: 3, val: 'A3' }, ...]
 */
function saveShiftCells(ym, changes) {
  requireAdmin_(); // 管理者だけが実行できる
  if (!ym) throw new Error('年月が指定されていません。');
  if (!changes || !changes.length) return { ok: true, saved: 0 };

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHIFT_SHEET);
  if (!sh) throw new Error('shift シートがありません。');

  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var last = sh.getLastRow();
    var rowOf = {};
    if (last >= 2) {
      var keys = sh.getRange(2, 1, last - 1, SHIFT_FIXED_COLS).getValues();
      for (var i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === ym) rowOf[String(keys[i][1]).trim()] = i + 2;
      }
    }

    var byNo = {};
    changes.forEach(function (c) {
      var no = String(c.no).trim();
      if (!byNo[no]) byNo[no] = [];
      byNo[no].push(c);
    });

    var nos = Object.keys(byNo).filter(function (no) {
      if (rowOf[no]) return true;
      // 行が無く、入れる値も全部空なら行を作らない
      return byNo[no].some(function (c) { return String(c.val == null ? '' : c.val) !== ''; });
    });
    nos.forEach(function (no) {
      if (!rowOf[no]) {
        var r = sh.getLastRow() + 1;
        if (r < 2) r = 2;
        sh.getRange(r, 1, 1, SHIFT_FIXED_COLS).setValues([[ym, no]]);
        rowOf[no] = r;
      }
    });

    nos.forEach(function (no) {
      var r = rowOf[no];
      var vals = sh.getRange(r, SHIFT_FIXED_COLS + 1, 1, SHIFT_MAX_DAYS).getValues()[0];
      byNo[no].forEach(function (c) {
        var d = Number(c.day);
        if (d >= 1 && d <= SHIFT_MAX_DAYS) vals[d - 1] = (c.val == null ? '' : String(c.val));
      });
      sh.getRange(r, SHIFT_FIXED_COLS + 1, 1, SHIFT_MAX_DAYS).setValues([vals]);
    });

    SpreadsheetApp.flush();
    return { ok: true, saved: changes.length };
  } finally {
    lock.releaseLock();
  }
}


/** 画面へは JSON 文字列で返す（null や NaN が混ざっても壊れないようにするため） */
function getShiftMonthJson(ym) {
  return JSON.stringify(getShiftMonth(ym));
}




/** ------------------------------------------------------------------
 *  申請（希望休・シフト希望）
 *  request シート: A=年月 / B=スタッフNo / C=日 / D=種別 / E=申請日 / F=入力者
 *  ------------------------------------------------------------------ */

var REQUEST_SHEET = 'request';
var REQ_KIND_OFF = '希望休';
var REQ_KIND_WORK = '出勤希望';
var REQ_KIND_PAID = '有給';

/** 【初回】request シートを作る */
function setupRequestSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(REQUEST_SHEET);
  var created = false;
  if (!sh) { sh = ss.insertSheet(REQUEST_SHEET); created = true; }

  var head = ['年月', 'スタッフNo', '日', '種別', '申請日', '入力者'];
  sh.getRange(1, 1, 1, head.length)
    .setValues([head])
    .setFontWeight('bold')
    .setBackground('#e8eaed');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(2, 90);
  sh.setColumnWidth(3, 50);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 200);
  sh.getRange(1, 1, sh.getMaxRows(), 4).setHorizontalAlignment('center');

  SpreadsheetApp.getUi().alert(created ? 'request シートを作りました。' : 'request シートの見出しを整えました。');
}

/** その月の申請を全部読む */
function getRequests_(ym) {
  var out = [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SHEET);
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, 6).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim() !== String(ym)) continue;
    var no = String(v[i][1]).trim();
    var day = Number(v[i][2]);
    if (!no || !day) continue;
    out.push({ no: no, day: day, kind: String(v[i][3]).trim(), at: v[i][4], by: String(v[i][5]).trim() });
  }
  return out;
}

/** 今操作している人のメールアドレス（取れなければ空文字） */
function currentUserEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}


/** ===== シフト希望の代行入力（2026/8/27 追加） ===== */
function openRequestDialog() {
  var t = HtmlService.createTemplateFromFile('request');
  t.nav = '';
  var html = t.evaluate()
    .setWidth(980).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, 'シフト希望の申請');
}

function reqDays_(ym) {
  var s = String(ym);
  var y = Number(s.slice(0, 4));
  var m = Number(s.slice(5, 7));
  var n = new Date(y, m, 0).getDate();
  var w = ['日', '月', '火', '水', '木', '金', '土'];
  var out = [];
  for (var d = 1; d <= n; d++) {
    out.push({ day: d, wd: w[new Date(y, m - 1, d).getDay()] });
  }
  return out;
}

function getRequestStaffJson() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('staff');
  if (!sh) return JSON.stringify([]);
  var last = sh.getLastRow();
  if (last < 2) return JSON.stringify([]);
  var v = sh.getRange(2, 1, last - 1, 29).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var no = String(v[i][0]).trim();
    if (!no) continue;
    var status = String(v[i][17]).trim();
    if (status && status !== '在籍') continue;
    out.push({
      no: no,
      name: String(v[i][1]).trim(),
      email: String(v[i][2]).trim(),
      floor: String(v[i][3]).trim(),
      subFloor: String(v[i][4]).trim(),
      ngDays: String(v[i][14]).trim(), // 稼働不可曜日（例「火,木」）。申請画面の注意に使う
      employment: String(v[i][5]).trim(),
      maxWish: (function(x){ var n = Number(x); return (x === '' || isNaN(n)) ? 0 : n; })(String(v[i][15]).trim()),
      paidLeft: (v[i][28] === '' || v[i][28] === null || v[i][28] === undefined) ? '' : v[i][28]
    });
  }
  return JSON.stringify(out);
}

function reqNormFloor_(s) {
  return normFloorCode_(s);
}

function reqSymbols_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('symbol');
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 4).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var code = String(v[i][0]).trim();
    if (!code) continue;
    out.push({
      code: code,
      kind: String(v[i][1]).trim(),
      floor: reqNormFloor_(v[i][2]),
      note: String(v[i][3]).trim()
    });
  }
  return out;
}

function getRequestMonthJson(ym, no) {
  var list = JSON.parse(getRequestStaffJson());
  var st = null;
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].no) === String(no)) st = list[i];
  }
  var marks = reqRowsFor_(ym, no);
  return JSON.stringify({
    ym: String(ym),
    staff: st,
    days: reqDays_(ym),
    marks: marks,
    symbols: reqSymbols_(),
    mainFloor: st ? reqNormFloor_(st.floor) : '',
    subFloor: st ? reqNormFloor_(st.subFloor) : '',
    kindOff: REQ_KIND_OFF,
    kindWork: REQ_KIND_WORK,
    kindPaid: REQ_KIND_PAID,
    paidLeft: st ? st.paidLeft : '',
    paidUsedOther: st ? paidLeaveCount_(paidLeaveIndex_(), String(no), String(ym), null) : 0,
    paidShiftDays: st ? paidShiftDays_(String(no), String(ym)) : [],
    me: currentUserEmail_()
  });
}

function saveRequests(ym, no, kindKey, items) {
  requireRequestRight_(no); // 管理者は全員分、ログイン済みの一般は自分の分だけ
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SHEET);
    if (!sh) throw new Error('requestシートがありません。メニューの「requestシートを作る」を先に実行してください。');
    if (String(sh.getRange(1, 7).getValue()).trim() !== '希望勤務') {
      sh.getRange(1, 7).setValue('希望勤務');
      sh.getRange(1, 7).setFontWeight('bold').setBackground('#e8eaed');
    }
    var kind = (kindKey === 'work') ? REQ_KIND_WORK : REQ_KIND_OFF;
    var last = sh.getLastRow();
    var keep = [];
    if (last >= 2) {
      var v = sh.getRange(2, 1, last - 1, 7).getValues();
      for (var i = 0; i < v.length; i++) {
        var a = String(v[i][0]).trim();
        var b = String(v[i][1]).trim();
        if (!a && !b) continue;
        if (reqYm_(v[i][0]) === String(ym) && b === String(no)) continue;
        keep.push([v[i][0], v[i][1], v[i][2], v[i][3], v[i][4], v[i][5], v[i][6]]);
      }
    }
    var now = new Date();
    var by = currentUserEmail_();
    var list = items || [];
    for (var d = 0; d < list.length; d++) {
      var k0 = (list[d].kind === 'paid') ? REQ_KIND_PAID : kind;
      keep.push([String(ym), String(no), Number(list[d].day), k0, now, by, String(list[d].val || '')]);
    }
    if (last >= 2) sh.getRange(2, 1, last - 1, 7).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 7).setValues(keep);
    return '保存しました（' + list.length + '日）';
  } finally {
    lock.releaseLock();
  }
}

function reqYm_(x) {
  if (x instanceof Date) {
    var mm = x.getMonth() + 1;
    return x.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm;
  }
  return String(x).trim();
}

function reqRowsFor_(ym, no) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SHEET);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 7).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (reqYm_(v[i][0]) !== String(ym)) continue;
    if (String(v[i][1]).trim() !== String(no)) continue;
    out.push({ day: Number(v[i][2]), kind: String(v[i][3]).trim(), val: String(v[i][6] === undefined ? "" : v[i][6]).trim() });
  }
  return out;
}


/** ===== 自動でシフトを組む（2026/8/28 追加・入れ物） ===== */
function openAutoBuildDialog() {
  var html = HtmlService.createTemplateFromFile('autobuild').evaluate()
    .setWidth(780).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'シフトを自動で組む');
}

function autoReadShift_(ym) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHIFT_SHEET);
  var map = {};
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  var v = sh.getRange(2, 1, last - 1, SHIFT_FIXED_COLS + SHIFT_MAX_DAYS).getValues();
  for (var i = 0; i < v.length; i++) {
    if (reqYm_(v[i][0]) !== String(ym)) continue;
    var no = String(v[i][1]).trim();
    if (!no) continue;
    var days = {};
    for (var d = 1; d <= SHIFT_MAX_DAYS; d++) {
      var s = String(v[i][SHIFT_FIXED_COLS + d - 1]).trim();
      if (s) days[d] = s;
    }
    map[no] = days;
  }
  return map;
}

function reqRowsAll_(ym) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_SHEET);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 7).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (reqYm_(v[i][0]) !== String(ym)) continue;
    var no = String(v[i][1]).trim();
    if (!no) continue;
    out.push({
      no: no,
      day: Number(v[i][2]),
      kind: String(v[i][3]).trim(),
      val: String(v[i][6] === undefined ? '' : v[i][6]).trim()
    });
  }
  return out;
}

function getAutoBuildInfoJson(ym) {
  var staff = readStaffForShift_();
  var cur = autoReadShift_(ym);
  var filled = 0, k, d;
  for (k in cur) { for (d in cur[k]) filled++; }
  var reqs = reqRowsAll_(ym);
  var off = 0, work = 0;
  for (var i = 0; i < reqs.length; i++) {
    if (reqs[i].kind === REQ_KIND_OFF || reqs[i].kind === REQ_KIND_PAID) off++; else work++;
  }
  return JSON.stringify({
    ym: String(ym),
    days: reqDays_(ym).length,
    staffCount: staff.length,
    filled: filled,
    reqOff: off,
    reqWork: work
  });
}

function runAutoBuild(ym, upto, mode, opt) {
  requireAdmin_(); // 管理者だけが実行できる
  opt = autoNormOpt_(opt);
  saveAutoBuildSettings_({ upto: upto, mode: mode, nightMax: opt.nightMax, needLeader: opt.needLeader, nightGap: opt.nightGap, maxRun: opt.maxRun });
  var log = [];
  var staff = readStaffForShift_();
  var cur = autoReadShift_(ym);
  var nDays = reqDays_(ym).length;
  var changes = [];
  var i, d, no, day;

  if (mode === 'clear') {
    var cleared = 0;
    for (i = 0; i < staff.length; i++) {
      var days = cur[staff[i].no] || {};
      for (d in days) { changes.push({ no: staff[i].no, day: Number(d), val: '' }); cleared++; }
      cur[staff[i].no] = {};
    }
    log.push('既存のシフトを消しました（' + cleared + 'マス）');
  } else {
    log.push('既に入っているマスはそのまま残します');
  }

  var reqs = reqRowsAll_(ym);
  var put0 = 0, skip0 = 0, putP = 0, skipP = 0;
  for (i = 0; i < reqs.length; i++) {
    var isOff0 = (reqs[i].kind === REQ_KIND_OFF), isPaid0 = (reqs[i].kind === REQ_KIND_PAID);
    if (!isOff0 && !isPaid0) continue;
    no = reqs[i].no;
    day = Number(reqs[i].day);
    if (!day || day > nDays) continue;
    if (!cur[no]) cur[no] = {};
    if (cur[no][day]) { if (isPaid0) skipP++; else skip0++; continue; }
    cur[no][day] = isPaid0 ? '有' : '希';
    changes.push({ no: no, day: day, val: isPaid0 ? '有' : '希' });
    if (isPaid0) putP++; else put0++;
  }
  log.push('ステップ0 希望休の反映：' + put0 + '件を「希」で置きました' + (skip0 ? '（' + skip0 + '件は既に入っていたので飛ばしました）' : ''));
  if (putP || skipP) log.push('ステップ0 有給の反映：' + putP + '件を「有」で置きました' + (skipP ? '（' + skipP + '件は既に入っていたので飛ばしました）' : ''));

  var order = ['step0', 'step1', 'step2', 'step3', 'all'];
  var goal = order.indexOf(upto);
  if (goal < 0) goal = 4;
  if (goal >= 1) autoBuildNight_(ym, cur, changes, nDays, opt, log);
  if (goal >= 2) autoBuildOff_(ym, cur, changes, nDays, opt, log);
  if (goal >= 3) autoBuildDay_(ym, cur, changes, nDays, opt, log);
  if (goal >= 4) autoBuildFree_(ym, cur, changes, nDays, opt, log);
  if (goal >= 4) autoFixRunLen_(ym, cur, changes, nDays, opt, log);

  autoSummary_(ym, cur, nDays, opt, log);

  if (changes.length) saveShiftCells(ym, changes);
  log.push('シフト表に書き込んだマス数：' + changes.length);
  return log.join(String.fromCharCode(10));
}


/** ------------------------------------------------------------------
 *  自動でシフトを組む — ステップ1 夜勤
 *  ------------------------------------------------------------------ */

var AUTO_SETTINGS_KEY = 'autoBuildSettings';
var AUTO_WD = ['日', '月', '火', '水', '木', '金', '土'];

/** 前回の設定を返す（画面の初期値に使う） */
function getAutoBuildSettingsJson() {
  var d = { upto: 'all', mode: 'keep', nightMax: 4, needLeader: true, nightGap: 3, maxRun: 6 };
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty(AUTO_SETTINGS_KEY);
    if (raw) {
      var o = JSON.parse(raw);
      if (o.upto) d.upto = String(o.upto);
      if (o.mode) d.mode = String(o.mode);
      var n = Number(o.nightMax);
      if (n && !isNaN(n)) d.nightMax = n;
      var gp = Number(o.nightGap);
      if (gp === 2 || gp === 3) d.nightGap = gp;
      var mr = Number(o.maxRun);
      if (mr && !isNaN(mr)) d.maxRun = mr;
      if (o.needLeader !== undefined && o.needLeader !== null) d.needLeader = (o.needLeader === true || String(o.needLeader) === 'true');
    }
  } catch (e) {}
  return JSON.stringify(d);
}

function saveAutoBuildSettings_(s) {
  try {
    PropertiesService.getDocumentProperties().setProperty(AUTO_SETTINGS_KEY, JSON.stringify(s));
  } catch (e) {}
}

function autoNormOpt_(opt) {
  var o = opt || {};
  var n = Number(o.nightMax);
  if (!n || isNaN(n) || n < 1) n = 4;
  var lead = true;
  if (o.needLeader !== undefined && o.needLeader !== null) lead = (o.needLeader === true || String(o.needLeader) === 'true');
  var gap = Number(o.nightGap);
  if (gap !== 2 && gap !== 3) gap = 3;
  var run = Number(o.maxRun);
  if (!run || isNaN(run) || run < 2) run = 6;
  return { nightMax: n, needLeader: lead, nightGap: gap, maxRun: run };
}

/** 「火,木」「火・木」などをばらして ['火','木'] にする */
function autoSplitList_(s) {
  var t = String(s || '');
  t = t.split('、').join(',').split('・').join(',').split('/').join(',');
  t = t.split(String.fromCharCode(32)).join(',').split('　').join(',');
  var a = t.split(',');
  var out = [];
  for (var i = 0; i < a.length; i++) {
    var x = a[i].trim();
    if (x) out.push(x.charAt(0));
  }
  return out;
}

function autoWeekday_(ym, day) {
  var p = String(ym).split('-');
  var dt = new Date(Number(p[0]), Number(p[1]) - 1, Number(day));
  return AUTO_WD[dt.getDay()];
}

/** 「3階南」→「S」のようにフロア名からコードを引く */
function autoFloorCode_(label) {
  var c = normFloorCode_(label);
  return (c && c !== '全' && FLOOR_LABEL[c]) ? c : '';
}

function autoNightSymbol_(code) {
  // symbol シートで「区分＝夜勤」かつそのフロア専用の記号があればそれを使う（例：新館→新夜）
  var list = getSymbolList_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].kind === '夜勤' && list[i].floor && normFloorCode_(list[i].floor) === code) return list[i].sym;
  }
  return '夜' + String.fromCharCode(32) + code;
}

/** 勤務パターンで「区分＝日勤」の対象になっているフロア（pattern シートの D 列と H 列） */
function patternDayFloors_() {
  var out = {};
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pattern');
    var v = sh ? sh.getDataRange().getValues() : [];
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][7] || '').trim() !== '日勤') continue;
      var t = String(v[i][3] || '').trim();
      if (t === '全') { floorCodes_().forEach(function (c) { out[c] = true; }); continue; }
      t.split(/[,、，]/).forEach(function (x) { var c = normFloorCode_(x); if (c) out[c] = true; });
    }
  } catch (e) {}
  return out;
}

function autoHeadSym_(v) {
  return String(v || '').split(String.fromCharCode(32)).join('　').split('　')[0];
}

function autoIsNight_(v) {
  var h = autoHeadSym_(v);
  if (h === '夜') return true;
  var m = getSymbolMap_()[h];
  return !!(m && m.kind === '夜勤');
}

/** staff シートを夜勤の判定に必要な項目つきで読む */
function autoStaffNight_(includeNotCounted) {
  var out = [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('staff');
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, 27).getValues();   // AA列（公休日数）まで読む
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    var no = String(r[0] || '').trim();
    if (!no) continue;
    var zaiseki = String(r[17] || '在籍').trim();
    if (zaiseki === '退職' || zaiseki === '休職中') continue;
    var job = String(r[18] || '').trim();
    if (job === '看護師' || job === '看護助手') continue;
    var counted = !(r[16] === true || String(r[16]) === 'TRUE');
    if (!counted && !includeNotCounted) continue;
    out.push({
      no: no,
      counted: counted,
      name: String(r[1] || no),
      floor: String(r[3] || '').trim(),
      sub: String(r[4] || '').trim(),
      night: (r[9] === true || String(r[9]) === 'TRUE'),
      senju: (r[10] === true || String(r[10]) === 'TRUE'),
      leader: (r[11] === true || String(r[11]) === 'TRUE'),
      nmax: autoNumOrZero_(r[21]),
      prio: autoPrioList_(r[22], r[23], r[24], r[25]),
      ngDays: autoSplitList_(String(r[14] || '')),
      emp: String(r[5] || '').trim(),
      offDays: r[26]
    });
  }
  return out;
}

function autoCanFloor_(st, code) {
  if (st.floor === '全') return true;
  if (st.floor === code) return true;
  var subs = autoSplitList_(st.sub);
  for (var i = 0; i < subs.length; i++) if (subs[i] === code) return true;
  return false;
}

/** 直近の夜勤が何日前か（7日より前なら 99） */
function autoLastNightGap_(cur, no, day) {
  var days = cur[no] || {};
  for (var k = 1; k <= 7; k++) {
    if (autoIsNight_(days[day - k])) return k;
  }
  return 99;
}

/** その日に夜勤を入れられるか */
function autoNightOk_(cur, st, day, nDays, ym, opt) {
  var days = cur[st.no] || {};
  if (days[day]) return false;
  if (day < nDays && days[day + 1]) return false;
  if (autoIsNight_(days[day - 1])) return false;
  if (opt && Number(opt.nightGap) >= 3 && autoIsNight_(days[day - 2])) return false;
  if (st.ngDays.length) {
    if (st.ngDays.indexOf(autoWeekday_(ym, day)) >= 0) return false;
    if (day < nDays && st.ngDays.indexOf(autoWeekday_(ym, day + 1)) >= 0) return false;
  }
  return true;
}

function autoNightScore_(st, cur, day, code, cnt, wish) {
  var sc = cnt[st.no] * 10;
  if (st.senju) sc -= 40;
  if (wish[st.no + '#' + day]) sc -= 25;
  if (st.floor !== code) sc += 6;
  var g = autoLastNightGap_(cur, st.no, day);
  if (g < 5) sc += (5 - g) * 4;
  return sc;
}

function autoPickNight_(cand, cur, day, code, cnt, wish, used, nDays, ym, opt, leaderOnly) {
  var best = null;
  for (var i = 0; i < cand.length; i++) {
    var st = cand[i];
    if (used[st.no]) continue;
    if (leaderOnly && !st.leader) continue;
    if (cnt[st.no] >= autoNightLimit_(st, opt, ym, nDays)) continue;
    if (!autoCanFloor_(st, code)) continue;
    if (!autoNightOk_(cur, st, day, nDays, ym, opt)) continue;
    var sc = autoNightScore_(st, cur, day, code, cnt, wish);
    if (best === null || sc < best.score) best = { st: st, score: sc };
  }
  return best;
}

function autoPutNight_(cur, changes, no, day, code, nDays, isLeader) {
  var sym = autoNightSymbol_(code);
  if (isLeader) sym = sym + '★';   // その日の夜勤リーダー
  if (!cur[no]) cur[no] = {};
  cur[no][day] = sym;
  changes.push({ no: no, day: day, val: sym });
  if (day < nDays && !cur[no][day + 1]) {
    cur[no][day + 1] = '明';
    changes.push({ no: no, day: day + 1, val: '明' });
  }
}

/** ステップ1 本体 */
function autoBuildNight_(ym, cur, changes, nDays, opt, log) {
  var all = autoStaffNight_();
  var cand = [];
  var i, d, no, dd;
  for (i = 0; i < all.length; i++) if (all[i].night) cand.push(all[i]);
  if (!cand.length) {
    log.push('ステップ1 夜勤：夜勤に入れる人がいません（スタッフマスタの夜勤可否を確認してください）');
    return;
  }

  var base = [];
  var stf = getStaffing_();
  for (i = 0; i < stf.length; i++) {
    if (String(stf[i].kind).trim() !== '夜勤') continue;
    var code = autoFloorCode_(stf[i].floor);
    if (!code) continue;
    var need = (stf[i].min === null ? 0 : Number(stf[i].min));
    var want = (stf[i].want === null ? need : Number(stf[i].want));
    if (need > 0) base.push({ code: code, must: true });
    for (var k = need; k < want; k++) base.push({ code: code, must: false });
  }
  if (!base.length) {
    log.push('ステップ1 夜勤：staffing シートに夜勤の行がありません');
    return;
  }

  // 最低人数の枠を先に、希望だけの枠（3階南など）は後から埋める
  var mustSlots = [], optSlots = [];
  for (i = 0; i < base.length; i++) { if (base[i].must) mustSlots.push(base[i]); else optSlots.push(base[i]); }
  base = mustSlots.concat(optSlots);

  var cnt = {}, wish = {};
  for (i = 0; i < cand.length; i++) cnt[cand[i].no] = 0;
  for (no in cur) {
    if (cnt[no] === undefined) continue;
    for (dd in cur[no]) if (autoIsNight_(cur[no][dd])) cnt[no]++;
  }
  var reqs = reqRowsAll_(ym);
  for (i = 0; i < reqs.length; i++) {
    if (reqs[i].kind !== REQ_KIND_WORK) continue;
    if (autoIsNight_(reqs[i].val)) wish[reqs[i].no + '#' + Number(reqs[i].day)] = true;
  }

  var senjuN = 0, leadN = 0, mustN = 0;
  for (i = 0; i < cand.length; i++) { if (cand[i].senju) senjuN++; if (cand[i].leader) leadN++; }
  for (i = 0; i < base.length; i++) if (base[i].must) mustN++;
  var per = Math.floor(nDays / opt.nightGap);
  var capacity = per * cand.length;
  // 人が足りない月は、日によって5人・0人と偏らないように1日あたりの人数をならす
  var dayCap = base.length;
  if (capacity < mustN * nDays) dayCap = Math.max(1, Math.floor(capacity / nDays));

  var placed = 0;
  var miss = [];
  for (d = 1; d <= nDays; d++) {
    var todo = [];
    var used = {};
    var filledToday = 0;
    // すでにシートに入っている夜勤（「残して埋める」のとき）は埋まり済みとして数える
    var have = {};
    for (i = 0; i < base.length; i++) {
      var c0 = base[i].code;
      if (have[c0] === undefined) have[c0] = autoNightSlotCount_(cur, d, c0);
      var done0 = have[c0] > 0;
      if (done0) { have[c0]--; filledToday++; }
      todo.push({ code: c0, must: base[i].must, filled: done0 });
    }
    // その日に入れる人が少ないフロアから先に埋める（新館などが後回しにならないように）
    for (i = 0; i < todo.length; i++) todo[i].cands = autoCountCand_(cand, cur, d, todo[i].code, cnt, used, nDays, ym, opt);
    todo.sort(function (a, b) { if (a.must !== b.must) return a.must ? -1 : 1; return a.cands - b.cands; });

    if (opt.needLeader && !autoDayHasLeader_(cur, cand, d)) {   // すでに★の人がいる日は飛ばす
      var best = null;
      for (i = 0; i < todo.length; i++) {
        if (!todo[i].must) continue;
        if (!floorLeaderOk_(todo[i].code)) continue;   // floor シートで夜勤リーダー不可のフロア
        var p = autoPickNight_(cand, cur, d, todo[i].code, cnt, wish, used, nDays, ym, opt, true);
        if (p) { best = { st: p.st, score: p.score, i: i }; break; }
      }
      if (best) {
        autoPutNight_(cur, changes, best.st.no, d, todo[best.i].code, nDays, true);
        todo[best.i].filled = true;
        used[best.st.no] = true;
        cnt[best.st.no]++;
        placed++;
        filledToday++;
      } else {
        miss.push(d + '日 リーダー');
      }
    }

    for (i = 0; i < todo.length; i++) {
      if (todo[i].filled) continue;
      if (filledToday >= dayCap) { if (todo[i].must) miss.push(d + '日 ' + (FLOOR_LABEL[todo[i].code] || todo[i].code)); continue; }
      var q = autoPickNight_(cand, cur, d, todo[i].code, cnt, wish, used, nDays, ym, opt, false);
      if (q) {
        autoPutNight_(cur, changes, q.st.no, d, todo[i].code, nDays);
        todo[i].filled = true;
        used[q.st.no] = true;
        cnt[q.st.no]++;
        placed++;
        filledToday++;
      } else if (todo[i].must) {
        miss.push(d + '日 ' + (FLOOR_LABEL[todo[i].code] || todo[i].code));
      }
    }
  }

    // ---- 2段階目 ----
  // 「明けの翌日は空ける」で組んだあと、まだ空いている必須の枠だけ
  // 「明けの翌日から入ってよい」でもう一度埋める（奥さんの回答⑩）
  var relaxed = 0;
  var relaxedList = [];
  if (Number(opt.nightGap) >= 3) {
    var opt2 = { nightMax: opt.nightMax, needLeader: opt.needLeader, nightGap: 2, maxRun: opt.maxRun };
    for (d = 1; d <= nDays; d++) {
      for (i = 0; i < base.length; i++) {
        if (!base[i].must) continue;
        if (autoNightSlotFilled_(cur, d, base[i].code)) continue;
        var used2 = {};
        var pick = null;
        var asLeader = false;
        if (opt.needLeader && floorLeaderOk_(base[i].code) && !autoDayHasLeader_(cur, cand, d)) {
          pick = autoPickNight_(cand, cur, d, base[i].code, cnt, wish, used2, nDays, ym, opt2, true);
          if (pick) asLeader = true;
        }
        if (!pick) pick = autoPickNight_(cand, cur, d, base[i].code, cnt, wish, used2, nDays, ym, opt2, false);
        if (pick) {
          autoPutNight_(cur, changes, pick.st.no, d, base[i].code, nDays, asLeader);
          cnt[pick.st.no]++;
          placed++;
          relaxed++;
          relaxedList.push(d + '日 ' + (FLOOR_LABEL[base[i].code] || base[i].code));
        }
      }
    }
    if (relaxed) {
      var rest = [];
      for (var mi = 0; mi < miss.length; mi++) {
        if (relaxedList.indexOf(miss[mi]) < 0) rest.push(miss[mi]);
      }
      miss = rest;
    }
  }
log.push('ステップ1 夜勤：' + placed + '件を置きました（翌日の「明」も自動で付けています）');
  if (relaxed) log.push('　うち ' + relaxed + '件は、人が足りないので明けの翌日に夜勤を入れて埋めました：' + relaxedList.slice(0, 25).join('／') + (relaxedList.length > 25 ? ' ほか' : ''));
  log.push('　夜勤に入れる人 ' + cand.length + '人（うち専従 ' + senjuN + '人・リーダー可 ' + leadN + '人）／1日の枠 最低' + mustN + '・希望' + base.length);
  log.push('　この月に必要な夜勤 最低' + (mustN * nDays) + '回。夜勤の間隔の設定では1人が月に最大' + per + '回、今の人数で出せるのは最大' + capacity + '回です');
  if (miss.length) {
    var head = miss.slice(0, 25).join('／');
    log.push('　埋まらなかったところ ' + miss.length + '件：' + head + (miss.length > 25 ? ' ほか' : ''));
  } else {
    log.push('　最低人数はすべて埋まりました');
  }
  var lines = [];
  for (i = 0; i < cand.length; i++) {
    if (cnt[cand[i].no]) lines.push(cand[i].name + String.fromCharCode(32) + cnt[cand[i].no] + '回');
  }
  if (lines.length) log.push('　1人あたりの夜勤回数：' + lines.join('／'));
}


/** その日そのフロアに入れる人が何人いるか数える */
function autoCountCand_(cand, cur, day, code, cnt, used, nDays, ym, opt) {
  var n = 0;
  for (var i = 0; i < cand.length; i++) {
    var st = cand[i];
    if (used[st.no]) continue;
    if (cnt[st.no] >= autoNightLimit_(st, opt, ym, nDays)) continue;
    if (!autoCanFloor_(st, code)) continue;
    if (!autoNightOk_(cur, st, day, nDays, ym, opt)) continue;
    n++;
  }
  return n;
}

/** シフト表をブラウザのタブで開く（全画面で見たいとき） */
/** 「別タブで開く」が開く管理者用ウェブアプリのURL（デプロイ「kanrisha-yo (fukushunin)」）。
 *  ScriptApp.getService().getUrl() はアーカイブ済みの古いデプロイのURLを返すことがあるので固定で持つ。
 *  管理者用のデプロイを作り直したときは、ここのURLも直すこと。 */
var ADMIN_WEBAPP_URL = '' /* ここに管理者用WebアプリのURLを入れる */;

function openShiftTableTab() {
  var ui = SpreadsheetApp.getUi();
  var url = ADMIN_WEBAPP_URL || '';
  if (!url) { try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; } }
  if (!url) {
    ui.alert('まだウェブアプリとして公開されていません。' + String.fromCharCode(10) + 'Apps Script の「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」で公開してください。');
    return;
  }
  var full = url + '?page=shift';
  var h = '<style>body{font-family:Meiryo,sans-serif;padding:16px;font-size:14px;line-height:1.9;color:#222}'
    + 'a{font-size:16px;color:#1a5fb4}p.note{color:#666;font-size:12px;line-height:1.7}</style>'
    + '<p>下のリンクを押すと、シフト表がブラウザのタブで開きます。画面いっぱいに表示できます。</p>'
    + '<p><a href="' + full + '" target="_blank" rel="noopener">シフト表を別のタブで開く</a></p>'
    + '<p class="note">開くのは「デプロイした版」です。コードを直したあとは、Apps Script の「デプロイ」→「デプロイを管理」→ 鉛筆マーク →「バージョン」を「新バージョン」にして更新しないと、古い画面が開きます。</p>';
  ui.showModalDialog(HtmlService.createHtmlOutput(h).setWidth(560).setHeight(260), 'シフト表を別タブで開く');
}

/** ------------------------------------------------------------------
 *  自動でシフトを組む — ステップ2 公休
 *  ------------------------------------------------------------------ */

/** 公休を置く対象のスタッフ（看護と退職・休職中は除く） */
function autoStaffForOff_() {
  var out = [];
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('staff');
  if (!sh) return out;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var v = sh.getRange(2, 1, last - 1, 29).getValues();
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    var no = String(r[0] || '').trim();
    if (!no) continue;
    var zai = String(r[17] || '在籍').trim();
    if (zai === '退職' || zai === '休職中') continue;
    var job = String(r[18] || '').trim();
    if (job === '看護師' || job === '看護助手') continue;
    out.push({
      no: no,
      name: String(r[1] || no),
      emp: String(r[5] || '').trim(),
      countable: !(r[16] === true || String(r[16]) === 'TRUE'),
      ngDays: autoSplitList_(String(r[14] || '')),
      leader: (r[11] === true || String(r[11]) === 'TRUE'),
      offDays: (r[26] === '' || r[26] === null || r[26] === undefined) ? '' : r[26],
      paidLeft: (r[28] === '' || r[28] === null || r[28] === undefined) ? '' : r[28]
    });
  }
  return out;
}

/** その人の公休の規定日数。staff シートの「公休日数」に数字が入っていればそれを使う */
function autoOffQuota_(st, quota) {
  var v = st ? st.offDays : '';
  if (v === '' || v === null || v === undefined) return quota;
  var n = Number(v);
  if (isNaN(n) || n < 0 || n > 31) return quota;
  return n;
}

/** 休みのマスか（公休・希望休。有休は別枠なので数えない） */
function autoIsRest_(val) {
  var h = autoHeadSym_(val);
  return (h === '休' || h === '希');
}

/** その日を含む連続勤務の長さ。空きマスはこれから勤務が入るものとして数える */
/** 連勤の切れ目になるマスか（公休・希望休・有休。有休は公休の枚数には数えないが、連勤は切れる） 2026/9/13 */
function autoIsBreak_(val) {
  return autoIsRest_(val) || autoHeadSym_(val) === '有';
}

function autoRunLen_(days, day, nDays) {
  var n = 1, k;
  for (k = day - 1; k >= 1; k--) { if (autoIsBreak_(days[k])) break; n++; }
  for (k = day + 1; k <= nDays; k++) { if (autoIsBreak_(days[k])) break; n++; }
  return n;
}

/** ステップ2 本体 */
function autoBuildOff_(ym, cur, changes, nDays, opt, log) {
  var all = autoStaffForOff_();
  var i, d, no, days;
  if (!all.length) { log.push('ステップ2 公休：対象のスタッフがいません'); return; }

  // パート希望日制：出勤希望が無い日を「休」にする
  var wishWork = {};
  var reqs = reqRowsAll_(ym);
  for (i = 0; i < reqs.length; i++) {
    if (reqs[i].kind !== REQ_KIND_WORK) continue;
    wishWork[reqs[i].no + '#' + Number(reqs[i].day)] = true;
  }
  var partPut = 0, partN = 0, hakenN = 0;
  for (i = 0; i < all.length; i++) {
    if (all[i].emp === '派遣・非常勤') { hakenN++; continue; }   // 公休はこのあと常勤と同じ扱いで置く
    if (all[i].emp !== 'パート希望日制') continue;
    partN++;
    no = all[i].no;
    if (!cur[no]) cur[no] = {};
    for (d = 1; d <= nDays; d++) {
      if (cur[no][d]) continue;
      if (wishWork[no + '#' + d]) continue;
      cur[no][d] = '休';
      changes.push({ no: no, day: d, val: '休' });
      partPut++;
    }
  }

  // 常勤：規定日数まで振る。派遣・非常勤も、休みの日数が決まるまでは常勤と同じ扱いにする
  var quota = requiredHolidayFor_(ym, nDays);
  var mem = [];
  for (i = 0; i < all.length; i++) if (all[i].emp === '常勤' || all[i].emp === '派遣・非常勤') mem.push(all[i]);

  // 1日に必要な出勤者数（夜勤は配置済みなので除く）
  var minNeed = 0;
  var stf = getStaffing_();
  for (i = 0; i < stf.length; i++) {
    if (String(stf[i].kind).trim() === '夜勤') continue;
    if (stf[i].min) minNeed += Number(stf[i].min);
  }

  // その日にあと何人休ませられるか
  var cap = [];
  for (d = 0; d <= nDays + 1; d++) cap[d] = 0;
  for (d = 1; d <= nDays; d++) {
    var free = 0;
    for (i = 0; i < all.length; i++) {
      if (!all[i].countable) continue;
      var dd = cur[all[i].no] || {};
      if (dd[d]) continue;
      free++;
    }
    cap[d] = Math.max(0, free - minNeed);
  }

  var need = {}, over = [], total = 0, indiv = [];
  for (i = 0; i < mem.length; i++) {
    no = mem[i].no;
    if (!cur[no]) cur[no] = {};
    var qi = autoOffQuota_(mem[i], quota);
    if (qi !== quota) indiv.push(mem[i].name + '（' + qi + '日）');
    var have = 0;
    for (d = 1; d <= nDays; d++) if (autoIsRest_(cur[no][d])) have++;
    var n = qi - have;
    if (n < 0) { over.push(mem[i].name + '（' + have + '日）'); n = 0; }
    need[no] = n;
    total += n;
  }

  var placed = 0, short = [], guard = 0, remain = {};
  while (total > 0 && guard < 5000) {
    guard++;
    var pick = null;
    for (i = 0; i < mem.length; i++) {
      if (need[mem[i].no] <= 0) continue;
      if (pick === null || need[mem[i].no] > need[pick.no]) pick = mem[i];
    }
    if (!pick) break;
    no = pick.no;
    days = cur[no];
    var bestD = 0, bestSc = -99999;
    for (d = 1; d <= nDays; d++) {
      if (days[d]) continue;
      if (cap[d] <= 0) continue;
      var run = autoRunLen_(days, d, nDays);
      var sc = run * 10;
      if (run > opt.maxRun) sc += 60;
      if (autoHeadSym_(days[d - 1]) === '明') sc += 25;
      if (pick.ngDays.length && pick.ngDays.indexOf(autoWeekday_(ym, d)) >= 0) sc += 15;
      if (autoIsBreak_(days[d - 1]) || autoIsBreak_(days[d + 1])) sc -= 12;
      if (sc > bestSc) { bestSc = sc; bestD = d; }
    }
    if (!bestD) {
      remain[no] = need[no];
      total -= need[no];
      need[no] = 0;
      continue;
    }
    days[bestD] = '休';
    changes.push({ no: no, day: bestD, val: '休' });
    cap[bestD]--;
    need[no]--;
    total--;
    placed++;
  }

  // 第2段階：人数の制限を外してでも規定日数まで置く（公休は所定の休みなので）
  var forced = 0, forcedDays = {};
  for (i = 0; i < mem.length; i++) {
    no = mem[i].no;
    var rest = remain[no] || 0;
    if (rest <= 0) continue;
    days = cur[no];
    while (rest > 0) {
      var bD = 0, bS = -99999;
      for (d = 1; d <= nDays; d++) {
        if (days[d]) continue;
        var run2 = autoRunLen_(days, d, nDays);
        var s2 = run2 * 10 + cap[d] * 5;
        if (run2 > opt.maxRun) s2 += 60;
        if (autoHeadSym_(days[d - 1]) === '明') s2 += 25;
        if (mem[i].ngDays.length && mem[i].ngDays.indexOf(autoWeekday_(ym, d)) >= 0) s2 += 15;
        if (autoIsBreak_(days[d - 1]) || autoIsBreak_(days[d + 1])) s2 -= 12;
        if (s2 > bS) { bS = s2; bD = d; }
      }
      if (!bD) { short.push(mem[i].name + '（あと' + rest + '日／空いているマスがありません）'); break; }
      days[bD] = '休';
      changes.push({ no: no, day: bD, val: '休' });
      cap[bD]--;
      forcedDays[bD] = (forcedDays[bD] || 0) + 1;
      forced++;
      placed++;
      rest--;
    }
  }

  log.push('ステップ2 公休：常勤と派遣・非常勤に' + placed + '件を置きました（規定' + quota + '日。公休日数を個別に決めている人はその日数／希望休は内数、有休は別枠）');
  if (partN) log.push('　パート希望日制 ' + partN + '人：出勤希望が無い日に「休」を' + partPut + '件置きました（日数のチェックはしません）');
  if (hakenN) log.push('　派遣・非常勤 ' + hakenN + '人：公休日数の欄が空いている人は、当面は常勤と同じ規定' + quota + '日で置いています');
  if (indiv.length) log.push('　公休日数を個別に決めている人：' + indiv.join('／'));
  if (over.length) log.push('　希望休が規定日数を超えている人：' + over.join('／'));
  if (forced) {
    var fl = [];
    for (d = 1; d <= nDays; d++) if (forcedDays[d]) fl.push(d + '日(' + forcedDays[d] + '人)');
    log.push('　人数が足りないので、最低人数を割ることを承知で' + forced + '件を置きました：' + fl.join('／'));
    log.push('　※公休は所定の休みなので規定日数を優先しました。上の日は出勤する人が足りません');
  }
  if (short.length) log.push('　規定日数に届かなかった人：' + short.join('／'));

}

/** 数字として読めれば数値、そうでなければ 0 を返す */
function autoNumOrZero_(x) {
  if (x === '' || x === null || x === undefined) return 0;
  var n = Number(x);
  if (isNaN(n) || n <= 0) return 0;
  return n;
}

/**
 * その人の「月の夜勤回数の上限」を返す。
 *  ・staff シートのV列に数字が入っていれば、それをその人の上限にする
 *  ・入っていなくて夜勤専従なら、上限なし（今までどおり）
 *  ・どちらでもなければ、自動生成の画面で入力した全員共通の値を使う
 */
function autoNightLimit_(st, opt, ym, nDays) {
  var base;
  if (st.nmax > 0) base = st.nmax;
  else if (st.senju) base = 999;
  else base = Number(opt.nightMax) || 0;
  // 公休を優先：夜勤1回で「夜・明」の2日を使うので、公休の規定日数が残る回数までにする
  // （パート希望日制は公休の決め方が違うので対象外）
  if (ym && nDays && st.emp !== 'パート希望日制') {
    var quota = autoOffQuota_(st, requiredHolidayFor_(ym, nDays));
    var cap = Math.floor((nDays - quota) / 2);
    if (cap < base) base = Math.max(0, cap);
  }
  return base;
}


/** ------------------------------------------------------------------
 *  その日の勤務者（掲示板と同じ並び）
 *  ------------------------------------------------------------------ */

/** メニューから開く */
function openDayBoard() {
  var t = HtmlService.createTemplateFromFile('dayboard');
  t.nav = '';
  SpreadsheetApp.getUi().showModalDialog(
    t.evaluate().setWidth(1200).setHeight(760), 'その日の勤務者');
}

/** フロアの言い方をコード（3・S・4・5・N）にそろえる */
function dayFloorCode_(x) {
  return normFloorCode_(x);
}

/** そのマスの人が、その日どのフロアにいるか */
function dayFloorOf_(val, st) {
  var s = String(val || '').replace('★', '');
  var sp = s.indexOf(String.fromCharCode(32));
  if (sp > 0) {
    var f = s.slice(sp + 1).trim();
    if (f) return dayFloorCode_(f);
  }
  var head = sp > 0 ? s.slice(0, sp) : s.trim();
  var mp = getSymbolMap_()[head];
  if (mp && mp.floor) return dayFloorCode_(mp.floor);
  return dayFloorCode_(st.floor);
}

/**
 * その日の勤務者を、掲示板と同じ形にまとめて返す。
 * ym は '2026-09' の形、day は 1〜31。
 */
function getDayBoardJson(ym, day) {
  var d = Number(day) || 1;
  var staff = readStaffForShift_();
  var rows = readShiftRows_(ym);

  var byNo = {};
  var i;
  for (i = 0; i < staff.length; i++) byNo[staff[i].no] = staff[i];

  // 夜勤リーダーに印を付けるための一覧
  var leaders = {};
  try {
    var ns = autoStaffNight_();
    for (i = 0; i < ns.length; i++) if (ns[i].leader) leaders[ns[i].no] = true;
  } catch (e) {}

  var dayFloors = patternDayFloors_();
  var floors = floorMaster_().map(function (f) { return { code: f.code, label: f.short || f.name, noDay: !dayFloors[f.code] }; });

  var out = { ym: ym, day: d, floors: floors, care: {}, kango: {}, free: [], night: [], after: [], etc: [] };
  for (i = 0; i < floors.length; i++) {
    out.care[floors[i].code] = { early: [], day: [], late: [] };
    out.kango[floors[i].code] = [];
  }

  for (var no in rows) {
    var st = byNo[no];
    if (!st) continue;
    var val = String(rows[no][d - 1] || '').trim();
    if (!val) continue;
    var isLeadCell = (val.indexOf('★') >= 0);
    if (isLeadCell) val = val.replace('★', '').trim();

    var kind = classifySymbol(val);
    if (kind === 'off' || kind === 'paid' || kind === 'wish') continue;

    var fl = dayFloorOf_(val, st);
    var item = { no: no, name: st.name || no, sym: val, floor: fl, kind: kind, leader: isLeadCell };

    if (kind === 'night') { out.night.push(item); continue; }
    if (kind === 'after') { out.after.push(item); continue; }
    if (kind === 'free')  { out.free.push(item); continue; }

    var isKango = (st.jobType === '看護師' || st.jobType === '看護助手');
    if (isKango) {
      if (out.kango[fl]) out.kango[fl].push(item); else out.etc.push(item);
      continue;
    }

    var slot = (kind === 'early') ? 'early' : (kind === 'late') ? 'late' : 'day';
    if (out.care[fl]) out.care[fl][slot].push(item); else out.etc.push(item);
  }

  return JSON.stringify(out);
}


/** その日、そのフロアの夜勤にすでに入っている人数（★付きも数える） */
function autoNightSlotCount_(cur, day, code) {
  var sym = autoNightSymbol_(code);
  var n = 0;
  for (var no in cur) {
    if (String(cur[no][day] || '').replace('★', '').trim() === sym) n++;
  }
  return n;
}

/** その日そのフロアの夜勤が、もう誰かで埋まっているか */
function autoNightSlotFilled_(cur, day, code) {
  return autoNightSlotCount_(cur, day, code) > 0;
}

/** その日、夜勤リーダーができる人がすでに夜勤に入っているか */
function autoDayHasLeader_(cur, cand, day) {
  for (var i = 0; i < cand.length; i++) {
    var days = cur[cand[i].no];
    if (days && String(days[day] || '').indexOf('★') >= 0) return true;
  }
  return false;
}


/** ------------------------------------------------------------------
 *  自動でシフトを組む — ステップ3 早番・遅番・日勤
 *  ------------------------------------------------------------------ */

/** 区分とフロアから、記号マスタの記号を引く（例 '早番' + '3' → 'A3'） */
function autoDaySymbol_(kindLabel, code) {
  var want = FLOOR_LABEL[code] || code;
  var list = getSymbolList_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].kind === kindLabel && list[i].floor === want) return list[i].sym;
  }
  return '';
}

/** その日その人を、この区分・フロアに入れられるか */
function autoDayOk_(cur, st, day, ym, code, kindLabel, sym) {
  var days = cur[st.no] || {};
  if (days[day]) return false;
  if (st.ngDays.length && st.ngDays.indexOf(autoWeekday_(ym, day)) >= 0) return false;
  if (!autoCanFloor_(st, code)) return false;
  if (st.prio.length && autoPrioRank_(st, kindLabel, sym) < 0) return false;   // 優先順位に書いていない区分・記号には入れない
  return true;
}

/** 選ぶ優先順位（数が小さい人を選ぶ） */
function autoDayScore_(st, cnt, wishKind, key) {
  var sc = (cnt[st.no] || 0) * 10;
  if (autoFloorCode_(st.floor) !== key.code) sc += 6;   // 主担当フロアの人を優先する
  var pi = autoPrioRank_(st, key.kind, key.sym);
  if (pi >= 0) sc -= (40 - pi * 10);   // 優先順位が上のものほど強く優先する
  if (wishKind === key.kind) sc -= 25;
  else if (wishKind) sc += 12;
  return sc;
}

/** その枠に入れる人が何人いるか */
function autoDayCount_(cand, cur, day, ym, code, used, kindLabel, sym) {
  var n = 0;
  for (var i = 0; i < cand.length; i++) {
    if (used[cand[i].no]) continue;
    if (autoDayOk_(cur, cand[i], day, ym, code, kindLabel, sym)) n++;
  }
  return n;
}

/** ステップ3 本体 */
function autoBuildDay_(ym, cur, changes, nDays, opt, log) {
  var cand = autoStaffNight_();
  var i, d, k;
  if (!cand.length) { log.push('ステップ3 早番・遅番・日勤：対象のスタッフがいません'); return; }

  // 出勤希望（no#日 → その人が出したい区分）
  var wish = {};
  var wishSym = {};
  var reqs = reqRowsAll_(ym);
  for (i = 0; i < reqs.length; i++) {
    if (reqs[i].kind !== REQ_KIND_WORK) continue;
    var sym = String(reqs[i].val || '').trim();
    var mp = getSymbolMap_()[autoHeadSym_(sym)];
    if (mp && mp.kind) {
      wish[reqs[i].no + '#' + Number(reqs[i].day)] = mp.kind;
      wishSym[reqs[i].no + '#' + Number(reqs[i].day)] = sym;
    }
  }

  // 枠を作る（最低ぶんを先、希望だけのぶんを後）
  var KINDS = ['早番', '遅番', '日勤'];
  var stf = getStaffing_();
  var base = [];
  for (i = 0; i < stf.length; i++) {
    if (KINDS.indexOf(stf[i].kind) < 0) continue;
    var code = autoFloorCode_(stf[i].floor);
    var mn = Number(stf[i].min) || 0;
    var wt = Number(stf[i].want) || 0;
    var sym2 = autoDaySymbol_(stf[i].kind, code);
    if (!sym2) continue;
    for (k = 0; k < mn; k++) base.push({ kind: stf[i].kind, code: code, sym: sym2, must: true });
    for (k = mn; k < wt; k++) base.push({ kind: stf[i].kind, code: code, sym: sym2, must: false });
  }
  if (!base.length) { log.push('ステップ3 早番・遅番・日勤：必要人数マスタに早番・遅番・日勤の行がありません'); return; }

  var cnt = {};
  var placed = 0;
  var wishPlaced = 0;
  var miss = [];

  for (d = 1; d <= nDays; d++) {
    var todo = [];
    for (i = 0; i < base.length; i++) {
      todo.push({ kind: base[i].kind, code: base[i].code, sym: base[i].sym, must: base[i].must });
    }
    var used = {};
    // 出勤希望（記号あり）は予約として先に置く（本人が出した希望なので、優先順位・稼働不可曜日より優先する）
    for (k = 0; k < cand.length; k++) {
      var wst = cand[k];
      var ws = wishSym[wst.no + '#' + d];
      if (!ws) continue;
      var wmp = getSymbolMap_()[autoHeadSym_(ws)];
      if (!wmp || KINDS.indexOf(wmp.kind) < 0) continue;
      if ((cur[wst.no] || {})[d]) continue;
      if (!cur[wst.no]) cur[wst.no] = {};
      cur[wst.no][d] = ws;
      changes.push({ no: wst.no, day: d, val: ws });
      used[wst.no] = true;
      cnt[wst.no] = (cnt[wst.no] || 0) + 1;
      placed++;
      wishPlaced++;
      for (i = 0; i < todo.length; i++) {
        if (todo[i].sym === ws) { todo.splice(i, 1); break; }
      }
    }
    // 入れる人が少ない枠から先に埋める
    for (i = 0; i < todo.length; i++) todo[i].cands = autoDayCount_(cand, cur, d, ym, todo[i].code, used, todo[i].kind, todo[i].sym);
    todo.sort(function (a, b) {
      if (a.must !== b.must) return a.must ? -1 : 1;
      return a.cands - b.cands;
    });

    for (i = 0; i < todo.length; i++) {
      var t = todo[i];
      var best = null, bestSc = 0;
      for (k = 0; k < cand.length; k++) {
        var st = cand[k];
        if (used[st.no]) continue;
        if (!autoDayOk_(cur, st, d, ym, t.code, t.kind, t.sym)) continue;
        var wk = wish[st.no + '#' + d] || '';
        var sc = autoDayScore_(st, cnt, wk, t);
        if (!best || sc < bestSc) { best = st; bestSc = sc; }
      }
      if (best) {
        if (!cur[best.no]) cur[best.no] = {};
        cur[best.no][d] = t.sym;
        changes.push({ no: best.no, day: d, val: t.sym });
        used[best.no] = true;
        cnt[best.no] = (cnt[best.no] || 0) + 1;
        placed++;
      } else if (t.must) {
        miss.push(d + '日 ' + t.kind + (FLOOR_LABEL[t.code] || t.code));
      }
    }
  }

  log.push('ステップ3 早番・遅番・日勤：' + placed + '件を置きました');
  if (wishPlaced) log.push('　出勤希望：' + wishPlaced + '件を希望の勤務で先に置きました（優先順位より優先）');
  if (miss.length) {
    log.push('　埋まらなかったところ ' + miss.length + '件：' + miss.slice(0, 25).join('／') + (miss.length > 25 ? ' ほか' : ''));
  } else {
    log.push('　最低人数はすべて埋まりました');
  }
}


/** 優先1〜4を配列にする（空欄は詰める） */
/**
 * 優先順位に、その区分・その記号が書かれているか。
 *  ・区分名（早番・遅番・日勤・F勤）で書かれていれば、その区分は全フロアOK
 *  ・記号（A3・B3・C5 など）で書かれていれば、その記号のフロアだけOK
 * 戻り値は書かれている位置（0が最優先）。書かれていなければ -1。
 */
function autoPrioRank_(st, kindLabel, sym) {
  if (!st.prio || !st.prio.length) return -1;
  var a = st.prio.indexOf(kindLabel);
  var b = sym ? st.prio.indexOf(sym) : -1;
  if (a < 0) return b;
  if (b < 0) return a;
  return (a < b) ? a : b;
}

function autoPrioList_(a, b, c, d) {
  var out = [];
  var v = [a, b, c, d];
  for (var i = 0; i < v.length; i++) {
    var t = String(v[i] || '').trim();
    if (t && out.indexOf(t) < 0) out.push(t);
  }
  return out;
}

/** ------------------------------------------------------------------
 *  自動でシフトを組む — ステップ4 F勤
 *  残っている空きマスを F にする。人数の上限は設けない。
 *  ------------------------------------------------------------------ */
function autoBuildFree_(ym, cur, changes, nDays, opt, log) {
  var cand = autoStaffNight_();
  if (!cand.length) { log.push('ステップ4 F勤：対象のスタッフがいません'); return; }

  var sym = autoDaySymbol_('F勤', '全');
  if (!sym) {
    var list = getSymbolList_();
    for (var k = 0; k < list.length; k++) if (list[k].kind === 'F勤') { sym = list[k].sym; break; }
  }
  if (!sym) { log.push('ステップ4 F勤：記号マスタに F勤 の記号がありません'); return; }

  var placed = 0;
  var skipped = 0;

  for (var i = 0; i < cand.length; i++) {
    var st = cand[i];
    // 優先順位に区分を書いていて、その中に F勤 が無い人は入れない
    if (st.prio.length && autoPrioRank_(st, 'F勤', sym) < 0) { skipped++; continue; }
    if (!cur[st.no]) cur[st.no] = {};
    for (var d = 1; d <= nDays; d++) {
      if (cur[st.no][d]) continue;
      if (st.ngDays.length && st.ngDays.indexOf(autoWeekday_(ym, d)) >= 0) continue;
      cur[st.no][d] = sym;
      changes.push({ no: st.no, day: d, val: sym });
      placed++;
    }
  }

  // 受け皿：優先順位に F勤 が無い人でも、行き場が無い空きマスは F勤 で埋める
  var extra = 0, extraNames = [];
  for (i = 0; i < cand.length; i++) {
    var st2 = cand[i];
    if (!(st2.prio.length && autoPrioRank_(st2, 'F勤', sym) < 0)) continue;
    if (!cur[st2.no]) cur[st2.no] = {};
    var n2 = 0;
    for (var d2 = 1; d2 <= nDays; d2++) {
      if (cur[st2.no][d2]) continue;
      if (st2.ngDays.length && st2.ngDays.indexOf(autoWeekday_(ym, d2)) >= 0) continue;
      cur[st2.no][d2] = sym;
      changes.push({ no: st2.no, day: d2, val: sym });
      n2++;
    }
    if (n2) { extra += n2; extraNames.push(st2.name + '（' + n2 + '日）'); }
  }

  // 3周目：人員カウント対象外の人（統括部長など）の空きマスも F勤 で埋める
  var outc = 0, outcNames = [];
  var allSt = autoStaffNight_(true);
  for (i = 0; i < allSt.length; i++) {
    var st3 = allSt[i];
    if (st3.counted) continue;
    if (!cur[st3.no]) cur[st3.no] = {};
    var n3 = 0;
    for (var d3 = 1; d3 <= nDays; d3++) {
      if (cur[st3.no][d3]) continue;
      if (st3.ngDays.length && st3.ngDays.indexOf(autoWeekday_(ym, d3)) >= 0) continue;
      cur[st3.no][d3] = sym;
      changes.push({ no: st3.no, day: d3, val: sym });
      n3++;
    }
    if (n3) { outc += n3; outcNames.push(st3.name + '（' + n3 + '日）'); }
  }

  log.push('ステップ4 F勤：' + (placed + extra + outc) + '件を「' + sym + '」で置きました（勤務も休みも決まっていなかった日を埋めました）');
  if (outc) log.push('　うち' + outc + '件は、人数に数えない人（統括部長など）の、決まっていなかった日を埋めたものです：' + outcNames.join('／'));
  if (extra) log.push('　うち' + extra + '件は、優先順位に F勤 が無い人の、行き場が無かった日を埋めたものです：' + extraNames.join('／'));
  if (skipped && !extra) log.push('　' + skipped + '人は優先順位に F勤 が入っていないので飛ばしました');
}

/** いちばん長い連勤（空きマス・休み・有休で切れる）を返す */
function autoLongestRun_(days, nDays) {
  var mx = 0, s = 0, e = 0, run = 0, st = 0;
  for (var d = 1; d <= nDays; d++) {
    var v = days[d];
    if (!v || autoIsRest_(v) || autoHeadSym_(v) === '有') { run = 0; continue; }
    if (run === 0) st = d;
    run++;
    if (run > mx) { mx = run; s = st; e = d; }
  }
  return { len: mx, start: s, end: e };
}

/**
 * 連勤の調整（全部置き終わったあとの2段階目）。
 * 連勤が上限を超えている人だけ、長い連勤の中の F勤 のマスと、
 * その人の公休の日を入れ替える。公休の枚数もF勤の本数も変わらず、
 * F勤は最低人数が0なので、その日に穴も開かない。
 */
function autoFixRunLen_(ym, cur, changes, nDays, opt, log) {
  var sym = autoDaySymbol_('F勤', '全');
  if (!sym) {
    var list = getSymbolList_();
    for (var k = 0; k < list.length; k++) if (list[k].kind === 'F勤') { sym = list[k].sym; break; }
  }
  if (!sym) return;

  var all = autoStaffNight_(true);
  var total = 0, names = [], gaveUp = [];
  for (var i = 0; i < all.length; i++) {
    var stf = all[i];
    var days = cur[stf.no];
    if (!days) continue;
    var moved = 0;
    for (var guard = 0; guard < 20; guard++) {
      var run = autoLongestRun_(days, nDays);
      if (run.len <= opt.maxRun) break;

      // 休みにする日：長い連勤の中の F勤 のマスで、できるだけ真ん中
      var mid = Math.floor((run.start + run.end) / 2);
      var d1 = 0, best = 9999, d;
      for (d = run.start; d <= run.end; d++) {
        if (autoHeadSym_(days[d]) !== sym) continue;
        var dist = Math.abs(d - mid);
        if (dist < best) { best = dist; d1 = d; }
      }
      if (!d1) break;

      // 勤務にする日：その人の公休の日。入れ替えて連勤が短くなる日だけ
      var d2 = 0;
      for (d = 1; d <= nDays; d++) {
        if (autoHeadSym_(days[d]) !== '休') continue;
        if (stf.ngDays.length && stf.ngDays.indexOf(autoWeekday_(ym, d)) >= 0) continue;
        var s1 = days[d1], s2 = days[d];
        days[d1] = '休'; days[d] = sym;
        if (autoLongestRun_(days, nDays).len < run.len) { d2 = d; break; }
        days[d1] = s1; days[d] = s2;
      }
      if (!d2) break;

      changes.push({ no: stf.no, day: d1, val: '休' });
      changes.push({ no: stf.no, day: d2, val: sym });
      moved++;
    }
    if (moved) { total += moved; names.push(stf.name + '（' + moved + '日）'); }
    else if (autoLongestRun_(days, nDays).len > opt.maxRun) gaveUp.push(stf.name);
  }

  if (total) log.push('連勤の調整：公休を' + total + '日ぶん動かして連勤を短くしました：' + names.join('／'));
  if (gaveUp.length) log.push('　動かせなかった人：' + gaveUp.join('／') + '（入れ替えられる F勤 か公休がありません）');
}
/** マスの値から「区分」と「フロア記号」を読み取る。休み・空きマスは null */
function autoCellKindFloor_(map, val) {
  val = String(val || '').replace('★', '');
  var head = autoHeadSym_(val);
  if (!head) return null;
  if (head === '休' || head === '希' || head === '有') return null;
  var mp = map[head];
  if (!mp || !mp.kind) return null;
  var sp = String(val).split(' ');
  var code = (sp.length > 1 && sp[1]) ? sp[1] : autoFloorCode_(mp.floor);
  return { kind: mp.kind, code: code };
}

/** 日の一覧を読みやすく（多いときは途中で切る） */
function autoDayListText_(arr) {
  var max = 12;
  var a = arr.slice(0, max).map(function (d) { return d + '日'; });
  return a.join('・') + (arr.length > max ? ' ほか' + (arr.length - max) + '日' : '');
}

/**
 * 要調整のまとめ。全部の工程が終わったあとに、人が判断すべきものだけを種類ごとに出す。
 * 自動では直せない（直すとどこかに穴が開く）ものばかりなので、副主任・部長が調整する材料にする。
 */
/**
 * 要調整の内容をデータで返す。ログ用の文章と画面の警告パネルの両方でこれを使う。
 *   lack      人が足りない枠 [{label, kind, floor, code, days:[日], short}]
 *   noLead    夜勤リーダーがいない日 [日]
 *   offShort  公休が規定日数に足りない人 [{no, name, days, quota}]
 *   runOver   連勤が上限を超えている人 [{no, name, run, start, end}]
 *   blank     まだ空いているマスがある人 [{no, name, days, dayList}]
 */
function shiftIssues_(ym, cur, nDays, opt) {
  var all = autoStaffForOff_();
  var map = getSymbolMap_();
  var i, d, k;
  opt = opt || {};
  var maxRun = Number(opt.maxRun) || 6;
  var needLeader = !!opt.needLeader;

  // その日・その区分・そのフロアに何人いるか
  var cnt = {};
  for (i = 0; i < all.length; i++) {
    if (!all[i].countable) continue;
    var days = cur[all[i].no] || {};
    for (d = 1; d <= nDays; d++) {
      var kf = autoCellKindFloor_(map, days[d]);
      if (!kf) continue;
      k = d + '#' + kf.kind + '#' + kf.code;
      cnt[k] = (cnt[k] || 0) + 1;
    }
  }

  // ① 人が足りない枠（最低人数に届いていない）
  var stf = getStaffing_();
  var lack = [], lackIx = {}, lackTotal = 0;
  for (i = 0; i < stf.length; i++) {
    var mn = Number(stf[i].min) || 0;
    if (!mn) continue;
    var code = autoFloorCode_(stf[i].floor);
    for (d = 1; d <= nDays; d++) {
      var have = cnt[d + '#' + stf[i].kind + '#' + code] || 0;
      if (have >= mn) continue;
      var label = stf[i].kind + ' ' + stf[i].floor;
      if (lackIx[label] === undefined) {
        lackIx[label] = lack.length;
        lack.push({ label: label, kind: stf[i].kind, floor: stf[i].floor, code: code, days: [], short: 0 });
      }
      lack[lackIx[label]].days.push(d);
      lack[lackIx[label]].short += (mn - have);
      lackTotal += (mn - have);
    }
  }

  // ② 夜勤リーダーがいない日
  var noLead = [], noLeadPick = [], noLeadNone = [], leaderNos = [];
  for (i = 0; i < all.length; i++) if (all[i].leader) leaderNos.push(all[i].no);
  if (needLeader) {
    for (d = 1; d <= nDays; d++) {
      var found = false, canPick = false;
      for (i = 0; i < all.length; i++) {
        var dz = cur[all[i].no] || {};
        var vv = String(dz[d] || '');
        if (vv.indexOf('★') >= 0) { found = true; break; }
        if (all[i].leader) {
          var kf = autoCellKindFloor_(map, vv);
          if (kf && kf.kind === '夜勤' && floorLeaderOk_(kf.code)) canPick = true;
        }
      }
      if (!found) {
        noLead.push(d);
        if (canPick) noLeadPick.push(d); else noLeadNone.push(d);
      }
    }
  }

  // ③ 公休が規定日数に届いていない人（常勤・派遣非常勤）
  var quota = requiredHolidayFor_(ym, nDays);
  var offShort = [];
  for (i = 0; i < all.length; i++) {
    if (all[i].emp !== '常勤' && all[i].emp !== '派遣・非常勤') continue;
    var dd = cur[all[i].no] || {};
    var n = 0;
    for (d = 1; d <= nDays; d++) if (autoIsRest_(dd[d])) n++;
    var qi = autoOffQuota_(all[i], quota);
    if (n < qi) offShort.push({ no: all[i].no, name: all[i].name, days: n, quota: qi });
  }

  // ④ 連勤が上限を超えている人
  var runOver = [];
  for (i = 0; i < all.length; i++) {
    var d3 = cur[all[i].no] || {};
    var r = autoLongestRun_(d3, nDays);
    if (r.len > maxRun) runOver.push({ no: all[i].no, name: all[i].name, run: r.len, start: r.start, end: r.end });
  }

  // ⑥ 有給が残日数を超えている人（申請＋シフト表の「有」を同じ日は1回として数える。残日数が空欄の人は見ない）
  var paidOver = [];
  var pidx = paidLeaveIndex_();
  for (i = 0; i < all.length; i++) {
    var pl = all[i].paidLeft;
    if (pl === '' || pl === null || pl === undefined || isNaN(Number(pl))) continue;
    var d5 = cur[all[i].no] || {}, curSet = {};
    for (d = 1; d <= nDays; d++) if (classifySymbol(d5[d]) === 'paid') curSet[d] = true;
    var used = paidLeaveCount_(pidx, all[i].no, String(ym), curSet);
    if (used > Number(pl)) paidOver.push({ no: all[i].no, name: all[i].name, used: used, left: Number(pl) });
  }

  // ⑤ まだ空いているマス
  var blank = [], filled = 0;
  for (i = 0; i < all.length; i++) {
    var d4 = cur[all[i].no] || {};
    var b = [];
    for (d = 1; d <= nDays; d++) if (!d4[d]) b.push(d);
    if (b.length) blank.push({ no: all[i].no, name: all[i].name, days: b.length, dayList: b });
    filled += (nDays - b.length);
  }

  return {
    ym: ym, nDays: nDays, maxRun: maxRun, needLeader: needLeader,
    lack: lack, lackTotal: lackTotal, noLead: noLead,
    offShort: offShort, runOver: runOver, blank: blank,
    paidOver: paidOver,
    noLeadPick: noLeadPick, noLeadNone: noLeadNone, leaderNos: leaderNos,
    staffCount: all.length, filled: filled,
    hasIssue: !!(lackTotal || noLead.length || offShort.length || runOver.length || blank.length || paidOver.length)
  };
}

/** 要調整のまとめをログに出す（中身は shiftIssues_ が数える） */
function autoSummary_(ym, cur, nDays, opt, log) {
  var r = shiftIssues_(ym, cur, nDays, opt);
  var i;

  log.push('──────── 要調整のまとめ ────────');

  if (r.lackTotal) {
    log.push('【人が足りない枠】' + r.lackTotal + '件');
    for (i = 0; i < r.lack.length; i++) {
      log.push('　' + r.lack[i].label + '：' + r.lack[i].days.length + '日　' + autoDayListText_(r.lack[i].days));
    }
  } else {
    log.push('【人が足りない枠】なし');
  }

  if (r.needLeader) {
    if (!r.noLead.length) {
      log.push('【夜勤リーダーがいない日】なし');
    } else {
      if (r.noLeadPick.length) log.push('【リーダーを決めていない日】' + r.noLeadPick.length + '日　' + autoDayListText_(r.noLeadPick) + '（その日の夜勤にリーダー可の人がいます。★を付けてください）');
      if (r.noLeadNone.length) log.push('【夜勤にリーダーできる人がいない日】' + r.noLeadNone.length + '日　' + autoDayListText_(r.noLeadNone) + '（夜勤の顔ぶれを入れ替える必要があります）');
    }
  }

  var t1 = [];
  for (i = 0; i < r.offShort.length; i++) t1.push(r.offShort[i].name + '（' + r.offShort[i].days + '日／規定' + r.offShort[i].quota + '日）');
  log.push(t1.length ? '【公休が規定日数に足りない人】' + t1.join('／') : '【公休が規定日数に足りない人】なし');

  var t2 = [];
  for (i = 0; i < r.runOver.length; i++) t2.push(r.runOver[i].name + '（' + r.runOver[i].run + '連勤）');
  log.push(t2.length ? '【' + r.maxRun + '連勤を超えている人】' + t2.join('／') : '【' + r.maxRun + '連勤を超えている人】なし');

  var t3 = [];
  for (i = 0; i < r.blank.length; i++) t3.push(r.blank[i].name + '（' + r.blank[i].days + '日）');
  log.push(t3.length ? '【勤務も休みも決まっていない日】' + t3.join('／') : '【勤務も休みも決まっていない日】なし');

  var t4 = [];
  for (i = 0; i < r.paidOver.length; i++) t4.push(r.paidOver[i].name + '（残' + r.paidOver[i].left + '日に対して' + r.paidOver[i].used + '日）');
  log.push(t4.length ? '【有給が残日数を超えている人】' + t4.join('／') : '【有給が残日数を超えている人】なし');

  if (r.hasIssue) {
    log.push('※ここから先は自動では直せません（直すとどこかに穴が開きます）。上の内容を見て調整してください');
  }
}

/**
 * シフト表の画面から呼ぶ。今シートに入っている内容のまま要調整の一覧を返す。
 * 自動生成を回さなくても、手で直したあとの状態で数え直せる。
 */
function getShiftIssues(ym, opt) {
  var nDays = daysInMonth_(ym);
  var rows = readShiftRows_(ym);
  var cur = {};
  Object.keys(rows).forEach(function (no) {
    var cells = rows[no], o = {};
    for (var d = 1; d <= nDays; d++) {
      var v = String(cells[d - 1] || '').trim();
      if (v) o[d] = v;
    }
    cur[no] = o;
  });
  opt = opt || {};
  var def = { maxRun: 6, needLeader: true };
  try { def = JSON.parse(getAutoBuildSettingsJson()); } catch (e) {}
  return shiftIssues_(ym, cur, nDays, {
    maxRun: Number(opt.maxRun) || Number(def.maxRun) || 6,
    needLeader: (opt.needLeader === undefined) ? (def.needLeader !== false) : !!opt.needLeader
  });
}

/** 画面から呼ぶ用。JSON文字列で返す */
function getShiftIssuesJson(ym, opt) {
  return JSON.stringify(getShiftIssues(ym, opt));
}


/** ===== 有給（2026/9/13 追加） =====
 * 残日数 ＝ staff シートの「有休残日数」 − 使った日数。
 * 使った日数は「request シートの有給の申請」と「shift シートの『有』」を、同じ人・同じ日は1回として数える（全部の月）。
 */
function paidLeaveIndex_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var req = {}, sh = {};
  var i, d, no, ym;
  var rs = ss.getSheetByName(REQUEST_SHEET);
  if (rs && rs.getLastRow() >= 2) {
    var rv = rs.getRange(2, 1, rs.getLastRow() - 1, 7).getValues();
    for (i = 0; i < rv.length; i++) {
      if (String(rv[i][3]).trim() !== REQ_KIND_PAID) continue;
      no = String(rv[i][1]).trim(); ym = reqYm_(rv[i][0]); d = Number(rv[i][2]);
      if (!no || !ym || !d) continue;
      if (!req[no]) req[no] = {};
      if (!req[no][ym]) req[no][ym] = {};
      req[no][ym][d] = true;
    }
  }
  var ts = ss.getSheetByName(SHIFT_SHEET);
  if (ts && ts.getLastRow() >= 2) {
    var tv = ts.getRange(2, 1, ts.getLastRow() - 1, SHIFT_FIXED_COLS + SHIFT_MAX_DAYS).getValues();
    for (i = 0; i < tv.length; i++) {
      no = String(tv[i][1]).trim(); ym = reqYm_(tv[i][0]);
      if (!no || !ym) continue;
      for (d = 1; d <= SHIFT_MAX_DAYS; d++) {
        var v = String(tv[i][SHIFT_FIXED_COLS + d - 1] || '').trim();
        if (!v || classifySymbol(v) !== 'paid') continue;
        if (!sh[no]) sh[no] = {};
        if (!sh[no][ym]) sh[no][ym] = {};
        sh[no][ym][d] = true;
      }
    }
  }
  return { req: req, sh: sh };
}

/**
 * 使った有給の日数を数える。
 * overYm を渡すと、その月のシフト表ぶんは overSet（{日:true}）で置き換える（自動生成の途中の状態で数えるため）。
 * overSet が null なら、その月は数えない（申請画面が自分で足すため）。
 */
function paidLeaveCount_(idx, no, overYm, overSet) {
  var req = idx.req[no] || {}, sh = idx.sh[no] || {};
  var months = {}, ym, d, n = 0;
  for (ym in req) months[ym] = true;
  for (ym in sh) months[ym] = true;
  if (overYm && overSet) months[overYm] = true;
  for (ym in months) {
    var set = {};
    if (overYm && ym === overYm) {
      if (!overSet) continue;
      for (d in (req[ym] || {})) set[d] = true;
      for (d in overSet) set[d] = true;
    } else {
      for (d in (req[ym] || {})) set[d] = true;
      for (d in (sh[ym] || {})) set[d] = true;
    }
    for (d in set) n++;
  }
  return n;
}

/** その月のシフト表で「有」が入っている日の一覧（申請画面で残を計算するときに使う） */
function paidShiftDays_(no, ym) {
  var idx = paidLeaveIndex_();
  var s = (idx.sh[no] || {})[ym] || {};
  var out = [];
  for (var d in s) out.push(Number(d));
  out.sort(function (a, b) { return a - b; });
  return out;
}
