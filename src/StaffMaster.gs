/**
 * 老健シフト管理アプリ
 * スタッフマスタ管理画面（第1弾）
 *
 * スプレッドシート「老健 シフトデータ」にコンテナバインドで配置する。
 * シート構成: floor / pattern / staff
 */

// ===== 定数 =====

var SHEET_STAFF = 'staff';
var SHEET_PATTERN = 'pattern';
var SHEET_FLOOR = 'floor';
var SHEET_QUAL = 'qual';
var SHEET_STAFFING = 'staffing';
var SHEET_SYMBOL = 'symbol';

// staff シートの列（1始まり）
var COL = {
  staffNo: 1,        // A スタッフNo
  name: 2,           // B 氏名
  email: 3,          // C メールアドレス
  floorMain: 4,      // D 担当フロア(主)
  floorSub: 5,       // E 担当フロア(兼務)
  employment: 6,     // F 雇用形態
  wage: 7,           // G 時給
  dependent: 8,      // H 扶養区分
  incomeYtd: 9,      // I 年初からの収入
  nightOk: 10,       // J 夜勤可否
  nightOnly: 11,     // K 夜勤専従
  nightLeader: 12,   // L 夜勤リーダー可否
  fStart: 13,        // M F勤開始時刻
  fEnd: 14,          // N F勤終了時刻
  ngDays: 15,        // O 稼働不可曜日
  wishLimit: 16,     // P 希望休上限
  notCounted: 17,    // Q 人員カウント対象外
  status: 18,        // R 在籍状況
  jobType: 19,       // S 職種
  group: 20,         // T 表示グループ（この画面では読むだけ）
  qual: 21,           // U 資格（カンマ区切り・複数可）
  nightMax: 22,       // V 夜勤回数の上限（空欄なら自動生成の画面の値を使う）
  prio1: 23,         // W 勤務の優先順位1
  prio2: 24,         // X 勤務の優先順位2
  prio3: 25,         // Y 勤務の優先順位3
  prio4: 26,          // Z 勤務の優先順位4
  offDays: 27,       // AA 公休日数（空欄なら規定どおり。派遣・非常勤など人ごとに違う場合に入れる）
  role: 28,          // AB 権限（「管理者」なら管理者、空欄なら一般）
  paidLeft: 29       // AC 有休残日数（管理者が付与時に書き換える。残＝この値−申請とシフト表の有給日数）
};

var LAST_COL = 19;   // 保存で書き込む範囲（A〜S）
var READ_COL = 29;   // 読み込む範囲（A〜AC）

// ===== Webアプリの入口 =====

function doGet(e) {
  var me = currentUser_();
  // URLに ?guest=1 を付けると、管理者でも一般スタッフの見え方を確認できる（権限は下がるだけ）
  var guest = !!(e && e.parameter && e.parameter.guest === '1');
  if (guest) me = { email: '', staffNo: '', name: '', admin: false };
  var page = (e && e.parameter && e.parameter.page) || (me.admin ? 'staff' : 'shift');
  // 管理者以外はスタッフマスタ・勤務パターン・申請の画面に入れない
  if (!me.admin && (page === 'staff' || page === 'pattern')) page = 'shift';
  // 申請は、管理者か、ログインしていてスタッフマスタに自分の行がある人だけ
  if (page === 'request' && !(me.admin || (me.email && me.staffNo))) page = 'shift';
  var file = (page === 'pattern') ? 'pattern' : (page === 'shift') ? 'shift' : (page === 'request') ? 'request' : (page === 'day') ? 'dayboard' : (page === 'feedback') ? 'feedback' : 'staff';
  var t = HtmlService.createTemplateFromFile(file);
  t.nav = navHtml_(file === 'dayboard' ? 'day' : file, me.admin, !!(me.email && me.staffNo), guest);
  return t
    .evaluate()
    .setTitle(appTitle_())
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===== 施設名（setting シート） =====
// setting シートの A列に「施設名」と書いた行の B列を施設名として使う。
// シートが無い・空欄のときは、これまでどおり「老健シフト管理」と表示する。
var SHEET_SETTING = 'setting';

/** 施設名を返す（無ければ空文字） */
function getFacilityName() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTING);
    if (!sh) return '';
    var v = sh.getDataRange().getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === '施設名') return String(v[i][1] || '').trim();
    }
  } catch (e) {}
  return '';
}

/** settingシートの値を項目名で読む（無ければ def を返す） */
function getSetting_(key, def) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTING);
    if (!sh) return def;
    var v = sh.getDataRange().getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === key) {
        var x = v[i][1];
        return (x === '' || x === null || x === undefined) ? def : x;
      }
    }
  } catch (e) {}
  return def;
}

/** その月の公休の規定日数。holiday シート（月｜公休日数）の値を使う。
    シートや値が無いときは、これまでどおり 30日以上の月は9日・それ以外は8日 */
var holidayCache_ = null;   // holiday シートを1回だけ読んで覚えておく（自動で組むときに何千回も呼ばれるため）

function requiredHolidayFor_(ym, days) {
  var def = days >= 30 ? 9 : 8;
  try {
    var month = Number(String(ym).split('-')[1]);
    if (!month) return def;
    if (holidayCache_ === null) {
      holidayCache_ = {};
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('holiday');
      var v = sh ? sh.getDataRange().getValues() : [];
      for (var i = 1; i < v.length; i++) {
        var m = Number(String(v[i][0]).replace('月', '').trim());
        var n = Number(v[i][1]);
        if (m && v[i][1] !== '' && !isNaN(n) && n >= 0 && n <= 31) holidayCache_[m] = n;
      }
    }
    return (holidayCache_[month] !== undefined) ? holidayCache_[month] : def;
  } catch (e) {}
  return def;
}

/* ------------------------------------------------------------------
 *  フロアマスタ（floor シート）
 *  A 記号(介護) / B 記号(看護) / C 名称 / D 並び順 / E 別名(カンマ区切り)
 *  F グループ(シフト表のタブ) / G 夜勤リーダー(チェック) / H 略称
 *  シートが無いときは、最初に作ったときの初期値を使う
 * ------------------------------------------------------------------ */
var FLOOR_DEFAULT_ = [
  { code: '3', kango: '3',  name: '3階',      order: 1, alias: ['3F', '三階'], group: '3階・3階南', leader: true,  short: '3F' },
  { code: 'S', kango: '南', name: '3階南',    order: 2, alias: ['南館', '3S', '3南', '南', '三階南'], group: '3階・3階南', leader: true, short: '3S' },
  { code: '4', kango: '4',  name: '4階',      order: 3, alias: ['4F', '四階'], group: '4階', leader: true, short: '4F' },
  { code: '5', kango: '56', name: '5階・6階', order: 4, alias: ['5階', '6階', '5・6階', '5階6階', '56'], group: '5階・6階', leader: true, short: '5・6F' },
  { code: 'N', kango: '新', name: '新館',     order: 5, alias: ['新'], group: '新館', leader: false, short: '新館' }
];
var floorCache_ = null;

/** フロアの一覧（並び順どおり） */
function floorMaster_() {
  if (floorCache_) return floorCache_;
  var list = [];
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('floor');
    if (sh) {
      var v = sh.getDataRange().getValues();
      for (var i = 1; i < v.length; i++) {
        var code = String(v[i][0] == null ? '' : v[i][0]).trim();
        if (!code) continue;
        var name = String(v[i][2] || '').trim() || code;
        var ord = Number(v[i][3]);
        var alias = String(v[i][4] || '').split(/[,、，]/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
        var lead = v[i][6];
        list.push({
          code: code,
          kango: String(v[i][1] == null ? '' : v[i][1]).trim(),
          name: name,
          order: isNaN(ord) || v[i][3] === '' ? 999 + i : ord,
          alias: alias,
          group: String(v[i][5] || '').trim() || name,
          leader: (lead === '' || lead === null || lead === undefined) ? true : (lead === true || String(lead).toUpperCase() === 'TRUE'),
          short: String(v[i][7] || '').trim() || name
        });
      }
    }
  } catch (e) {}
  if (!list.length) list = FLOOR_DEFAULT_.slice();
  list.sort(function (a, b) { return a.order - b.order; });
  floorCache_ = list;
  return list;
}

/** 記号 → 名称 の対応（'全' は 全フロア） */
function floorLabelMap_() {
  var m = {};
  floorMaster_().forEach(function (f) { m[f.code] = f.name; });
  m['全'] = '全フロア';
  return m;
}

/** 記号の並び */
function floorCodes_() {
  return floorMaster_().map(function (f) { return f.code; });
}

/** シフト表のグループ（統括部長 → 各フロアのグループ → フリー） */
function floorGroups_() {
  var out = ['統括部長'];
  floorMaster_().forEach(function (f) { if (out.indexOf(f.group) < 0) out.push(f.group); });
  if (out.indexOf('フリー') < 0) out.push('フリー');
  return out;
}

/** いろいろな書き方のフロアを記号にそろえる（見つからなければそのまま返す） */
function normFloorCode_(s) {
  var t = String(s == null ? '' : s).trim();
  if (!t) return '';
  var u = t.split(' ').join('').split('　').join('');
  if (['全', '全フロア', '施設全体', '全体', '共通'].indexOf(u) >= 0) return '全';
  var list = floorMaster_();
  for (var i = 0; i < list.length; i++) {
    var f = list[i];
    if (u === f.code || u === f.name || u === f.kango || u === f.short || f.alias.indexOf(u) >= 0) return f.code;
    if (u.toUpperCase() === String(f.code).toUpperCase()) return f.code;
  }
  return t;
}

/** そのフロアの夜勤からリーダーを選べるか */
function floorLeaderOk_(code) {
  var list = floorMaster_();
  for (var i = 0; i < list.length; i++) if (list[i].code === code) return list[i].leader;
  return true;
}

/** 画面（HTML）に渡すフロア情報 */
function floorInfoJson_() {
  return JSON.stringify(floorMaster_());
}

/** 画面の上に出すアプリ名（例：サンプル施設　シフト管理） */
function appTitle_() {
  var f = getFacilityName();
  return f ? f + '　シフト管理' : '老健シフト管理';
}

/** HTML に埋め込むときの文字の置き換え */
function escFacility_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** HTML から他のファイルを読み込むためのヘルパー */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** スプレッドシートのメニューにも入口を作る */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('シフト管理')
    .addItem('スタッフマスタを開く', 'showStaffDialog')
    .addItem('勤務パターンマスタを開く', 'showPatternDialog')
    .addSeparator()
    .addItem('シフト表を開く', 'openShiftTable')
    .addItem('シフト表を別タブで開く（全画面）', 'openShiftTableTab')
    .addItem('シフト希望の申請', 'openRequestDialog')
    .addItem('質問・要望', 'openFeedbackDialog')
    .addItem('その日の勤務者を見る', 'openDayBoard')
    .addItem('シフトを自動で組む', 'openAutoBuildDialog')
    .addSeparator()
    .addItem('【初回】shiftシートを作る', 'setupShiftSheet')
    .addItem('【初回】requestシートを作る', 'setupRequestSheet')
    .addItem('【確認用】サンプルシフトを入れる', 'setupShiftSampleData')
    .addItem('【確認用】サンプルスタッフを入れる', 'setupStaffSampleData')
    .addItem('【確認用】サンプルスタッフ40人に作り直す', 'setupSampleStaff')
    .addItem('【確認用】夜勤のサンプル設定を入れる', 'setupNightSampleData')
    .addSeparator()
    .addItem('【メンテ】マスタの初期設定', 'setupStatusAndQual')
    .addItem('【メンテ】夜勤回数の上限の列を追加', 'setupNightMaxColumn')
    .addItem('【メンテ】勤務の優先順位の列を追加・選択肢を更新', 'setupPrioColumns')
    .addSeparator()
    .addItem('【検証用】自分のアドレスを表示', 'whoAmI')
    .addItem('【検証用】自分の権限を表示', 'showMyRole')
    .addToUi();
}

function showStaffDialog() {
  var t = HtmlService.createTemplateFromFile('staff');
  t.nav = '';
  var html = t
    .evaluate()
    .setWidth(1100)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'スタッフマスタ');
}

// ===== 共通 =====

function getSheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('シートが見つかりません: ' + name);
  return sh;
}

/** 時刻セルの値を "H:mm" の文字列にする */
function toTimeString_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'H:mm');
  }
  return String(v);
}

/** "9:30" のような文字列を、スプレッドシートに書ける形にする */
function fromTimeString_(s) {
  if (!s) return '';
  return String(s).trim();
}

function toBool_(v) {
  return v === true || v === 'TRUE' || v === 'true' || v === '可' || v === '対象外';
}

// ===== マスタの選択肢 =====

/** 画面のプルダウンに出す選択肢をまとめて返す */
function getMasterOptions() {
  var floors = [];
  var fsh = getSheet_(SHEET_FLOOR);
  var fvals = fsh.getDataRange().getValues();
  for (var i = 1; i < fvals.length; i++) {
    var code = String(fvals[i][0] || '').trim();
    var label = String(fvals[i][2] || '').trim();
    if (code) floors.push({ code: code, label: label || code, group: String(fvals[i][5] || '').trim() || label || code });
  }

  return {
    floors: floors,
    jobTypes: ['介護職員', '看護師', '看護助手'],
    employments: ['常勤', 'パート希望日制', '派遣・非常勤'],
    dependents: ['103万', '106万', '130万', '150万'],
    statuses: ['在籍', '休職中', '退職'],
    groups: floorGroups_(),
    quals: getQuals_(),
    prios: getPrioOptions_(),
    weekdays: ['月', '火', '水', '木', '金', '土', '日']
  };
}

/**
 * 勤務の優先順位で選べる値の一覧。
 * 区分名（早番・遅番・日勤・F勤＝そのフロア全部OK）と、
 * 記号マスタの記号（A3・B3・C5 など＝そのフロアだけOK）を並べて返す。
 */
function getPrioOptions_() {
  var KINDS = ['早番', '遅番', '日勤', 'F勤'];
  var out = [];
  for (var i = 0; i < KINDS.length; i++) out.push({ code: KINDS[i], label: '全フロア' });
  var list = [];
  try { list = getSymbolList_(); } catch (e) { list = []; }
  for (i = 0; i < list.length; i++) {
    var s = list[i];
    if (KINDS.indexOf(s.kind) < 0) continue;
    var sym = String(s.sym || '').trim();
    if (!sym) continue;
    var dup = false;
    for (var j = 0; j < out.length; j++) if (out[j].code === sym) dup = true;
    if (dup) continue;
    out.push({ code: sym, label: s.kind + (s.floor ? ' ' + s.floor : '') });
  }
  return out;
}

/** qual シートから資格の一覧を読む（並び順はシートの行順） */
function getQuals_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_QUAL);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] || '').trim();
    if (v) out.push(v);
  }
  return out;
}

// ===== スタッフマスタ =====

/** 一覧用にスタッフを全件返す */
function getStaffList() {
  var sh = getSheet_(SHEET_STAFF);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var values = sh.getRange(2, 1, last - 1, READ_COL).getValues();
  var list = [];

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!String(r[COL.staffNo - 1] || '').trim() && !String(r[COL.name - 1] || '').trim()) {
      continue; // 空行は飛ばす
    }
    list.push(rowToStaff_(r, i + 2));
  }
  return list;
}

function rowToStaff_(r, rowIndex) {
  return {
    row: rowIndex,
    staffNo: String(r[COL.staffNo - 1] || ''),
    name: String(r[COL.name - 1] || ''),
    email: String(r[COL.email - 1] || ''),
    floorMain: String(r[COL.floorMain - 1] || ''),
    floorSub: String(r[COL.floorSub - 1] || ''),
    employment: String(r[COL.employment - 1] || ''),
    wage: r[COL.wage - 1] === '' ? '' : r[COL.wage - 1],
    dependent: String(r[COL.dependent - 1] || ''),
    incomeYtd: r[COL.incomeYtd - 1] === '' ? '' : r[COL.incomeYtd - 1],
    nightOk: toBool_(r[COL.nightOk - 1]),
    nightOnly: toBool_(r[COL.nightOnly - 1]),
    nightLeader: toBool_(r[COL.nightLeader - 1]),
    nightMax: r[COL.nightMax - 1] === '' || r[COL.nightMax - 1] === null || r[COL.nightMax - 1] === undefined ? '' : r[COL.nightMax - 1],
    prio1: String(r[COL.prio1 - 1] || ''),
    prio2: String(r[COL.prio2 - 1] || ''),
    prio3: String(r[COL.prio3 - 1] || ''),
    prio4: String(r[COL.prio4 - 1] || ''),
    fStart: toTimeString_(r[COL.fStart - 1]),
    fEnd: toTimeString_(r[COL.fEnd - 1]),
    ngDays: String(r[COL.ngDays - 1] || ''),
    wishLimit: r[COL.wishLimit - 1] === '' ? 3 : r[COL.wishLimit - 1],
    notCounted: toBool_(r[COL.notCounted - 1]),
    group: String(r[COL.group - 1] || ''),
    qual: String(r[COL.qual - 1] || ''),
    offDays: (r[COL.offDays - 1] === '' || r[COL.offDays - 1] === null || r[COL.offDays - 1] === undefined) ? '' : r[COL.offDays - 1],
    role: String(r[COL.role - 1] || ''),
    paidLeft: (r[COL.paidLeft - 1] === '' || r[COL.paidLeft - 1] === null || r[COL.paidLeft - 1] === undefined) ? '' : r[COL.paidLeft - 1],
    status: String(r[COL.status - 1] || '在籍'),
    jobType: String(r[COL.jobType - 1] || '介護職員')
  };
}

/**
 * スタッフを保存する。
 * s.row が 0 のときは新規追加、それ以外はその行を更新。
 * 戻り値は保存後の一覧。
 */
function saveStaff(s) {
  requireAdmin_(); // 管理者だけが実行できる
  var sh = getSheet_(SHEET_STAFF);

  var err = validateStaff_(s, sh);
  if (err) throw new Error(err);

  var row = Number(s.row) || 0;
  if (!row) {
    row = sh.getLastRow() + 1;
    if (row < 2) row = 2;
  }

  var values = [
    s.staffNo,
    s.name,
    s.email,
    s.floorMain,
    s.floorSub,
    s.employment,
    isPart_(s.employment) ? numOrBlank_(s.wage) : '',
    isPart_(s.employment) ? s.dependent : '',
    isPart_(s.employment) ? numOrBlank_(s.incomeYtd) : '',
    !!s.nightOk,
    !!s.nightOnly,
    !!s.nightLeader,
    fromTimeString_(s.fStart),
    fromTimeString_(s.fEnd),
    s.ngDays,
    numOrBlank_(s.wishLimit),
    !!s.notCounted,
    s.status || '在籍',
    s.jobType || '介護職員'
  ];

  sh.getRange(row, 1, 1, LAST_COL).setValues([values]);
  sh.getRange(row, COL.group).setValue(s.group || '');
  sh.getRange(row, COL.qual).setValue(s.qual || '');
  sh.getRange(row, COL.nightMax).setValue(numOrBlank_(s.nightMax));
  sh.getRange(row, COL.offDays).setValue(numOrBlank_(s.offDays));
  sh.getRange(row, COL.role).setValue(s.role || '');
  sh.getRange(row, COL.paidLeft).setValue(numOrBlank_(s.paidLeft));
  sh.getRange(row, COL.prio1, 1, 4).setValues([[s.prio1 || '', s.prio2 || '', s.prio3 || '', s.prio4 || '']]);
  SpreadsheetApp.flush();
  return getStaffList();
}

function isPart_(employment) {
  return employment === 'パート希望日制';
}

function numOrBlank_(v) {
  if (v === '' || v === null || v === undefined) return '';
  var n = Number(v);
  return isNaN(n) ? '' : n;
}

/** 入力チェック。問題があればメッセージを返す */
function validateStaff_(s, sh) {
  if (!String(s.name || '').trim()) return '氏名を入力してください。';
  if (!String(s.staffNo || '').trim()) return 'スタッフNoを入力してください。';
  if (!String(s.floorMain || '').trim()) return '担当フロア(主)を選んでください。';
  if (!String(s.employment || '').trim()) return '雇用形態を選んでください。';
  if (!String(s.jobType || '').trim()) return '職種を選んでください。';
  if (!String(s.group || '').trim()) return 'シフト表の表示グループを選んでください。';

  if (s.offDays !== '' && s.offDays !== null && s.offDays !== undefined) {
    var od = Number(s.offDays);
    if (isNaN(od) || od < 0 || od > 31) return '公休日数は 0〜31 の数字で入力してください。';
  }

  if (s.paidLeft !== '' && s.paidLeft !== null && s.paidLeft !== undefined) {
    var pl = Number(s.paidLeft);
    if (isNaN(pl) || pl < 0 || pl > 99) return '有休残日数は 0〜99 の数字で入力してください。';
  }

  if (s.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s.email).trim())) {
    return 'メールアドレスの形式が正しくありません。';
  }

  // スタッフNo の重複チェック（自分自身は除く）
  var last = sh.getLastRow();
  if (last >= 2) {
    var nos = sh.getRange(2, COL.staffNo, last - 1, 1).getValues();
    for (var i = 0; i < nos.length; i++) {
      var rowIndex = i + 2;
      if (rowIndex === Number(s.row)) continue;
      if (String(nos[i][0] || '').trim().toUpperCase() === String(s.staffNo).trim().toUpperCase()) {
        return 'スタッフNo「' + s.staffNo + '」は既に使われています。';
      }
    }
  }

  // 主担当と兼務が同じフロアになっていないか
  if (s.floorSub) {
    var subs = String(s.floorSub).split(',').map(function (x) { return x.trim(); });
    if (subs.indexOf(String(s.floorMain).trim()) >= 0) {
      return '兼務フロアに主担当と同じフロアが入っています。';
    }
  }

  // F勤の時間は両方そろっているか
  var hasStart = !!String(s.fStart || '').trim();
  var hasEnd = !!String(s.fEnd || '').trim();
  if (hasStart !== hasEnd) {
    return 'F勤の勤務時間は、開始と終了の両方を入力してください。';
  }

  var limit = Number(s.wishLimit);
  if (s.wishLimit !== '' && (isNaN(limit) || limit < 0 || limit > 31)) {
    return '希望休の目安は0〜31の数字で入力してください。';
  }

  if (isPart_(s.employment)) {
    if (s.wage !== '' && isNaN(Number(s.wage))) return '時給は数字で入力してください。';
    if (s.incomeYtd !== '' && isNaN(Number(s.incomeYtd))) return '年初からの収入は数字で入力してください。';
  }

  return null;
}

/**
 * スタッフを1件完全に消す。
 * 2026/8/24 時点では画面から呼んでいません（削除ボタンは廃止し、退職者は在籍状況＝退職で残す運用）。
 * 誤登録を消したいときはスプレッドシートの staff シートを直接編集します。
 */
function deleteStaff(row) {
  requireAdmin_(); // 管理者だけが実行できる
  var r = Number(row);
  if (!r || r < 2) throw new Error('削除する行が正しくありません。');
  var sh = getSheet_(SHEET_STAFF);
  sh.deleteRow(r);
  SpreadsheetApp.flush();
  return getStaffList();
}


/** 【初回だけ実行】在籍状況のプルダウンに「休職中」を足し、U列に「資格」の見出しを付ける */
function setupStatusAndQual() {
  var sh = getSheet_(SHEET_STAFF);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['在籍', '休職中', '退職'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, COL.status, 199, 1).setDataValidation(rule);
  sh.getRange(1, COL.qual).setValue('資格').setFontWeight('bold').setBackground('#e8eaed');

  // qual シート（資格マスタ）が無ければ作る
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qs = ss.getSheetByName(SHEET_QUAL);
  if (!qs) {
    qs = ss.insertSheet(SHEET_QUAL);
    qs.getRange(1, 1, 1, 2).setValues([['資格名', '備考']]);
    qs.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8eaed');
    qs.setFrozenRows(1);
    var rows = [['正看護師', ''], ['准看護師', ''], ['介護福祉士', ''], ['実務者研修修了', ''], ['初任者研修修了', ''], ['認知症基礎研修修了', ''], ['ヘルパー２級', '初任者研修修了の旧名称']];
    qs.getRange(2, 1, rows.length, 2).setValues(rows);
    qs.autoResizeColumns(1, 2);
  }
  setupStaffingSheet();
  setupSymbolSheet();
  SpreadsheetApp.flush();
  Logger.log('done');
}


/** 【初回だけ実行】pattern シートに「区分」列を足し、staffing（必要人数）シートを作る */
function setupStaffingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- pattern シートに区分列（H列）を用意する ---
  var KUBUN_OF = { 'A': '早番', 'A前倒し': '早番', 'B': '遅番', 'C': '日勤', 'F9': 'F勤', 'F': 'F勤', '夜勤': '夜勤', '新夜': '夜勤' };
  var KUBUN_LIST = ['早番', '遅番', '日勤', 'F勤', '夜勤'];
  var psh = getSheet_(SHEET_PATTERN);
  psh.getRange(1, 8).setValue('区分').setFontWeight('bold').setBackground('#e8eaed');
  var plast = psh.getLastRow();
  if (plast >= 2) {
    var names = psh.getRange(2, 1, plast - 1, 1).getValues();
    var kubun = psh.getRange(2, 8, plast - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      var nm = String(names[i][0] || '').trim();
      if (!String(kubun[i][0] || '').trim() && KUBUN_OF[nm]) kubun[i][0] = KUBUN_OF[nm];
    }
    psh.getRange(2, 8, plast - 1, 1).setValues(kubun);
    var prule = SpreadsheetApp.newDataValidation().requireValueInList(KUBUN_LIST, true).setAllowInvalid(false).build();
    psh.getRange(2, 8, 199, 1).setDataValidation(prule);
  }
  psh.autoResizeColumns(8, 1);

  // --- staffing シート（必要人数）を作る ---
  var sh = ss.getSheetByName(SHEET_STAFFING);
  if (sh) return; // 既にあれば何もしない（入力済みの値を消さないため）
  sh = ss.insertSheet(SHEET_STAFFING);
  sh.getRange(1, 1, 1, 4).setValues([['区分', 'フロア', '最低', '希望']]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);

  var order = floorCodes_();
  var floorNames = [];
  for (var j = 0; j < order.length; j++) floorNames.push(FLOOR_LABEL[order[j]]);

  var rows = [];
  ['早番', '遅番', '日勤', 'F勤', '夜勤'].forEach(function (k) {
    floorNames.forEach(function (f) {
      var v = (k === '早番' || k === '遅番') ? 1 : '';
      rows.push([k, f, v, v]);
    });
  });
  sh.getRange(2, 1, rows.length, 4).setValues(rows);

  var kr = SpreadsheetApp.newDataValidation().requireValueInList(KUBUN_LIST, true).setAllowInvalid(false).build();
  sh.getRange(2, 1, 199, 1).setDataValidation(kr);
  var fr = SpreadsheetApp.newDataValidation().requireValueInList(floorNames.concat(['施設全体']), true).setAllowInvalid(false).build();
  sh.getRange(2, 2, 199, 1).setDataValidation(fr);
  sh.autoResizeColumns(1, 4);
}


/** 【初回だけ実行】symbol（記号マスタ）シートを作る */
function setupSymbolSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(SHEET_SYMBOL)) return; // 既にあれば触らない
  var sh = ss.insertSheet(SHEET_SYMBOL);
  sh.getRange(1, 1, 1, 4).setValues([['記号', '区分', 'フロア', '説明']]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);

  var rows = [
    ['休', '公休', '', '公休'],
    ['有', '有休', '', '有給休暇'],
    ['希', '希望休', '', '希望休'],
    ['明', '明け', '', '夜勤明け（〜9:30）'],
    ['夜', '夜勤', '', '夜勤（16:30〜）'],
    ['新夜', '夜勤', '新館', '新館の夜勤'],
    ['A3', '早番', '3階', '7:30-16:30'],
    ['AS', '早番', '3階南', '7:30-16:30'],
    ['A4', '早番', '4階', '7:30-16:30'],
    ['A5', '早番', '5階・6階', '7:30-16:30'],
    ['AN', '早番', '新館', '9:00-17:00'],
    ['B3', '遅番', '3階', '11:00-20:00'],
    ['BS', '遅番', '3階南', '11:00-20:00'],
    ['B4', '遅番', '4階', '11:00-20:00'],
    ['C5', '日勤', '5階・6階', '8:30-17:30'],
    ['CS', '日勤', '3階南', '8:30-17:30'],
    ['CN', '日勤', '新館', '8:30-17:30'],
    ['F8', '日勤', '', '8:30-17:30'],
    ['F9', 'F勤', '', '9:00-18:00'],
    ['F', 'F勤', '', '9:00-18:00'],
    ['A', '看護', '', '看護 9:30-16:30'],
    ['B', '看護', '', '看護 8:00-17:00'],
    ['C', '看護', '', '看護 8:30-17:30'],
    ['D', '看護', '', '看護 9:00-16:00'],
    ['E', '看護', '', '看護 8:30-16:30'],
    ['3R', '看護', '3階', '看護リーダー'],
    ['新NS', '看護', '新館', '新館の看護']
  ];
  sh.getRange(2, 1, rows.length, 4).setValues(rows);

  var KUBUN = ['早番', '遅番', '日勤', 'F勤', '夜勤', '明け', '公休', '有休', '希望休', '看護'];
  var kr = SpreadsheetApp.newDataValidation().requireValueInList(KUBUN, true).setAllowInvalid(false).build();
  sh.getRange(2, 2, 199, 1).setDataValidation(kr);

  var order = floorCodes_();
  var floorNames = [];
  for (var j = 0; j < order.length; j++) floorNames.push(FLOOR_LABEL[order[j]]);
  var fr = SpreadsheetApp.newDataValidation().requireValueInList(floorNames, true).setAllowInvalid(true).build();
  sh.getRange(2, 3, 199, 1).setDataValidation(fr);

  sh.autoResizeColumns(1, 4);
}


/** 【確認用】いろいろな条件のスタッフをまとめて入れる（S004以降。既にあれば何もしない） */
function setupStaffSampleData() {
  var sh = getSheet_(SHEET_STAFF);
  var last = sh.getLastRow();
  if (last >= 2) {
    var nos = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < nos.length; i++) {
      if (String(nos[i][0]).trim() === 'S004') {
        SpreadsheetApp.getUi().alert('サンプルスタッフはすでに入っています。');
        return;
      }
    }
  }

  // A スタッフNo / B 氏名 / C メール / D 主 / E 兼務 / F 雇用形態 / G 時給 / H 扶養 / I 年収
  // J 夜勤可 / K 夜勤専従 / L 夜勤L / M F開始 / N F終了 / O 不可曜日 / P 希望休 / Q 対象外 / R 在籍 / S 職種 / T 表示グループ / U 資格
  var rows = [
    ['S004','山田 一郎','s004@example.com','3','S','常勤','','','',true,false,true,'','','',3,false,'在籍','介護職員','3階・3階南','介護福祉士'],
    ['S005','佐藤 花子','s005@example.com','3','','パート希望日制',1100,'103万',300000,false,false,false,'','','火,金',3,false,'在籍','介護職員','3階・3階南','実務者研修修了'],
    ['S006','鈴木 次郎','s006@example.com','S','3','常勤','','','',true,true,false,'','','',3,false,'在籍','介護職員','3階・3階南','初任者研修修了'],
    ['S007','高橋 みどり','s007@example.com','3','','常勤','','','',false,false,false,'','','',3,false,'在籍','看護師','3階・3階南','正看護師'],
    ['S008','田中 春子','s008@example.com','4','','常勤','','','',true,false,true,'','','',3,false,'在籍','介護職員','4階','介護福祉士'],
    ['S009','伊藤 健','s009@example.com','4','','派遣・非常勤','','','',true,false,false,'','','',3,false,'在籍','介護職員','4階','初任者研修修了'],
    ['S010','渡辺 良子','s010@example.com','4','','常勤','','','',false,false,false,'','','',3,false,'休職中','看護助手','4階',''],
    ['S011','中村 大輔','s011@example.com','5','','常勤','','','',true,false,true,'','','',3,false,'在籍','介護職員','5階・6階','介護福祉士'],
    ['S012','小林 さゆり','s012@example.com','5','N','パート希望日制',1150,'130万',420000,false,false,false,'','','水',3,false,'在籍','介護職員','5階・6階','実務者研修修了'],
    ['S013','加藤 京子','s013@example.com','5','','常勤','','','',false,false,false,'','','',3,false,'在籍','看護師','5階・6階','准看護師'],
    ['S014','吉田 学','s014@example.com','N','','常勤','','','',true,false,true,'','','',3,false,'在籍','介護職員','新館','介護福祉士'],
    ['S015','山本 恵','s015@example.com','N','','派遣・非常勤','','','',true,true,false,'','','',3,false,'在籍','介護職員','新館','ヘルパー２級'],
    ['S016','松本 千鶴','s016@example.com','N','','常勤','','','',false,false,false,'','','',3,false,'在籍','看護師','新館','正看護師'],
    ['S017','井上 拓也','s017@example.com','全','','常勤','','','',false,false,false,'9:00','18:00','',3,false,'在籍','介護職員','フリー','介護福祉士'],
    ['S018','木村 直美','s018@example.com','全','','パート希望日制',1100,'106万',250000,false,false,false,'9:00','18:00','月',3,false,'在籍','介護職員','フリー','認知症基礎研修修了'],
    ['S019','清水 保','s019@example.com','3','','常勤','','','',false,false,false,'','','',3,false,'退職','介護職員','3階・3階南','介護福祉士'],
    ['S020','森 由紀','s020@example.com','4','','常勤','','','',false,false,false,'','','',3,false,'在籍','看護助手','4階','']
  ];

  var start = Math.max(sh.getLastRow() + 1, 2);
  sh.getRange(start, 1, rows.length, 21).setValues(rows);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('サンプルスタッフを ' + rows.length + ' 名入れました。');
}


/** ------------------------------------------------------------------
 *  【確認用】夜勤のサンプル設定
 *  自動でシフトを組む機能のテスト用に、夜勤に入れる人を増やす。
 *  さわるのは E列(兼務フロア)・J列(夜勤可否)・L列(夜勤リーダー可否) の3つだけ。
 *  ------------------------------------------------------------------ */
function setupNightSampleData() {
  var ui = SpreadsheetApp.getUi();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('staff');
  if (!sh) { ui.alert('staff シートがありません。'); return; }
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('スタッフのデータがありません。'); return; }

  var v = sh.getRange(2, 1, last - 1, 21).getValues();
  var FLOORS = floorCodes_();
  var rows = [];
  var i, f;

  for (i = 0; i < v.length; i++) {
    var r = v[i];
    var no = String(r[0] || '').trim();
    if (!no) continue;
    var zai = String(r[17] || '在籍').trim();
    if (zai === '退職' || zai === '休職中') continue;
    var job = String(r[18] || '').trim();
    if (job === '看護師' || job === '看護助手') continue;
    if (r[16] === true || String(r[16]) === 'TRUE') continue;
    var emp = String(r[5] || '').trim();
    if (emp === 'パート希望日制') continue;
    rows.push({
      row: i + 2,
      no: no,
      main: String(r[3] || '').trim(),
      sub: String(r[4] || '').trim(),
      leader: (r[11] === true || String(r[11]) === 'TRUE')
    });
  }
  if (!rows.length) { ui.alert('対象になるスタッフがいません。'); return; }

  // 今のフロアごとの人数を数える
  var cover = {};
  for (i = 0; i < FLOORS.length; i++) cover[FLOORS[i]] = 0;
  for (i = 0; i < rows.length; i++) {
    var mn = rows[i].main;
    if (mn === '全') { for (f = 0; f < FLOORS.length; f++) cover[FLOORS[f]]++; }
    else if (cover[mn] !== undefined) cover[mn]++;
    var sb = String(rows[i].sub).split('、').join(',').split('・').join(',').split(' ').join(',').split(',');
    for (f = 0; f < sb.length; f++) { var c = sb[f].trim().charAt(0); if (c && cover[c] !== undefined) cover[c]++; }
  }

  // 兼務が空の人に、いちばん手薄なフロアを1つ入れる
  var setSub = 0;
  for (i = 0; i < rows.length; i++) {
    if (rows[i].sub) continue;
    if (rows[i].main === '全') continue;
    var pick = '', lowest = -1;
    for (f = 0; f < FLOORS.length; f++) {
      var fl = FLOORS[f];
      if (fl === rows[i].main) continue;
      if (lowest < 0 || cover[fl] < lowest) { lowest = cover[fl]; pick = fl; }
    }
    if (!pick) continue;
    sh.getRange(rows[i].row, 5).setValue(pick);
    cover[pick]++;
    rows[i].sub = pick;
    setSub++;
  }

  // 全員を夜勤可にする
  for (i = 0; i < rows.length; i++) sh.getRange(rows[i].row, 10).setValue(true);

  // リーダー可を8人以上にする
  var lead = 0;
  for (i = 0; i < rows.length; i++) if (rows[i].leader) lead++;
  var addLead = 0;
  for (i = 0; i < rows.length && lead < 8; i++) {
    if (rows[i].leader) continue;
    sh.getRange(rows[i].row, 12).setValue(true);
    lead++;
    addLead++;
  }

  var msg = '夜勤に入れる人を ' + rows.length + '人にしました。' + String.fromCharCode(10)
    + '兼務フロアを入れた人：' + setSub + '人' + String.fromCharCode(10)
    + 'リーダー可にした人：' + addLead + '人（合計 ' + lead + '人）' + String.fromCharCode(10)
    + 'フロアごとの人数：';
  for (f = 0; f < FLOORS.length; f++) msg += ' ' + FLOOR_LABEL[FLOORS[f]] + ' ' + cover[FLOORS[f]] + '人';
  ui.alert(msg);
}

/** ウェブアプリの上に出す共通のナビゲーション（ダイアログでは出さない） */
function navHtml_(page, isAdmin, canRequest, guest) {
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }
  if (!url) return '';
  var items = [
    { key: 'staff', label: 'スタッフマスタ' },
    { key: 'pattern', label: '勤務パターン' },
    { key: 'shift', label: 'シフト表' },
    { key: 'day', label: 'その日の勤務者' },
    { key: 'request', label: 'シフト希望の申請' },
    { key: 'feedback', label: '質問・要望' }
  ];
  if (!isAdmin) items = items.filter(function (x) { return x.key === 'shift' || x.key === 'day' || x.key === 'feedback' || (canRequest && x.key === 'request'); });
  var h = '<div class="gnav"><span class="gnav-title">' + escFacility_(appTitle_()) + '</span>';
  for (var i = 0; i < items.length; i++) {
    var on = (items[i].key === page) ? ' gnav-on' : '';
    h += '<a class="gnav-a' + on + '" target="_top" href="' + url + '?page=' + items[i].key + (guest ? '&guest=1' : '') + '">' + items[i].label + '</a>';
  }
  h += '</div>';
  h += '<style>'
    + '.gnav{display:flex;align-items:center;gap:4px;background:#1B5E4A;padding:9px 16px;border-radius:6px;margin:0 0 14px;font-family:"Meiryo","Hiragino Kaku Gothic ProN",sans-serif;flex-wrap:wrap}'
    + '.gnav-title{color:#fff;font-size:15px;font-weight:bold;margin-right:18px}'
    + '.gnav-a{color:#dff0e8;text-decoration:none;font-size:14px;padding:6px 14px;border-radius:4px}'
    + '.gnav-a:hover{background:#2c7a62;color:#fff}'
    + '.gnav-on{background:#fff;color:#1B5E4A;font-weight:bold}'
    + '</style>';
  return h;
}

/**
 * staff シートの V列に「夜勤回数の上限」の見出しを入れる。
 * 既に入っていれば何もしない。データ行にはいっさい触れない。
 */
function setupNightMaxColumn() {
  var sh = getSheet_(SHEET_STAFF);
  var cell = sh.getRange(1, COL.nightMax);
  var now = String(cell.getValue() || '').trim();

  if (now === '夜勤回数の上限') {
    SpreadsheetApp.getUi().alert('すでに追加されています。');
    return;
  }
  if (now !== '') {
    SpreadsheetApp.getUi().alert('V1 に別の値（' + now + '）が入っています。念のため中止しました。');
    return;
  }

  cell.setValue('夜勤回数の上限');
  cell.setFontWeight('bold').setBackground('#e8eaed');
  sh.autoResizeColumn(COL.nightMax);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('staff シートの V列に「夜勤回数の上限」を追加しました。');
}


/**
 * staff シートの W〜Z列に「優先1〜優先4」の見出しを入れる。
 * 既に入っていれば何もしない。データ行にはいっさい触れない。
 */
function setupPrioColumns() {
  var sh = getSheet_(SHEET_STAFF);
  var head = ['優先1', '優先2', '優先3', '優先4'];
  var rng = sh.getRange(1, COL.prio1, 1, 4);
  var now = rng.getValues()[0];

  var already = true;
  var dirty = '';
  for (var i = 0; i < 4; i++) {
    var t = String(now[i] || '').trim();
    if (t !== head[i]) already = false;
    if (t !== '' && t !== head[i]) dirty = t;
  }
  if (dirty) {
    SpreadsheetApp.getUi().alert('W1〜Z1 に別の値（' + dirty + '）が入っています。念のため中止しました。');
    return;
  }

  if (!already) {
    rng.setValues([head]);
    rng.setFontWeight('bold').setBackground('#e8eaed');
  }

  // 選べる値をプルダウンにする（区分名＋記号マスタの記号）
  var opts = getPrioOptions_();
  var vals = [];
  for (var k = 0; k < opts.length; k++) vals.push(opts[k].code);
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(false).build();
  sh.getRange(2, COL.prio1, 300, 4).setDataValidation(rule);

  sh.autoResizeColumns(COL.prio1, 4);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(already
    ? '優先1〜優先4のプルダウンを最新にしました（区分名＋記号）。'
    : 'staff シートの W〜Z列に「優先1〜優先4」を追加しました。');
}

// ===== 権限（管理者かどうかの判定） =====

/**
 * いまアプリを使っている人が誰かを返す。
 * 匿名で開ける一般用URLではメールアドレスが取れないので email は空になる。
 */
function currentUser_() {
  var email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { email = ''; }
  var me = { email: email, staffNo: '', name: '', admin: false };
  if (!email) return me;

  // スプレッドシートのオーナーは常に管理者（全員の権限を消しても入れなくならないように）
  try {
    var owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
    if (owner && String(owner.getEmail() || '').trim().toLowerCase() === email) me.admin = true;
  } catch (e2) {}

  var sh = getSheet_(SHEET_STAFF);
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, READ_COL).getValues();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (String(r[COL.email - 1] || '').trim().toLowerCase() !== email) continue;
      me.staffNo = String(r[COL.staffNo - 1] || '');
      me.name = String(r[COL.name - 1] || '');
      if (String(r[COL.role - 1] || '').trim() === '管理者') me.admin = true;
      break;
    }
  }
  return me;
}

/** 画面から呼ぶ用 */
function getCurrentUserJson(asGuest) {
  if (asGuest) return JSON.stringify({ email: '', staffNo: '', name: '', admin: false });
  return JSON.stringify(currentUser_());
}

/** 書き換える操作の先頭で呼ぶ。管理者でなければ止める */
function requireAdmin_() {
  var me = currentUser_();
  if (!me.admin) throw new Error('この操作には管理者の権限が必要です。');
  return me;
}

/** 申請を保存してよいか。管理者は全員分、ログイン済みの一般は自分の分だけ */
function requireRequestRight_(staffNo) {
  var me = currentUser_();
  if (me.admin) return me;
  if (me.email && me.staffNo && String(staffNo) === me.staffNo) return me;
  throw new Error('自分の分だけ申請できます。（ほかの人の分は副主任にお願いしてください）');
}

/** 【検証用】いまの自分の権限を表示する */
function showMyRole() {
  var me = currentUser_();
  var msg = 'メールアドレス：' + (me.email || '（取れませんでした）') + '\n'
    + '氏名：' + (me.name || '（スタッフマスタに見つかりません）') + '\n'
    + 'スタッフNo：' + (me.staffNo || '－') + '\n'
    + '権限：' + (me.admin ? '管理者' : '一般');
  SpreadsheetApp.getUi().alert('自分の権限', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
