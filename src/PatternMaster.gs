/**
 * 老健シフト管理アプリ
 * 勤務パターンマスタ管理画面（第2弾）
 *
 * pattern シートを読み書きする。共通関数（getSheet_ / toTimeString_ /
 * fromTimeString_ / toBool_ / include）は StaffMaster.gs にある。
 */

// ===== 定数 =====

// pattern シートの列（1始まり）
var PCOL = {
  name: 1,       // A パターン名
  start: 2,      // B 開始時刻
  end: 3,        // C 終了時刻
  floors: 4,     // D 対象フロア
  desc: 5,       // E 説明
  overnight: 6,  // F 日をまたぐ
  perStaff: 7    // G 勤務時間はスタッフごと
};

var PLAST_COL = 7;

// 「勤務時間はスタッフごと」のときに時刻欄へ入れる文字
var PER_STAFF_LABEL = '個別';

// 対象フロアが全フロアのときの文字
var ALL_FLOORS_LABEL = '全';

// ===== 画面を開く =====

function showPatternDialog() {
  var t = HtmlService.createTemplateFromFile('pattern');
  t.nav = '';
  var html = t
    .evaluate()
    .setWidth(1100)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '勤務パターンマスタ');
}

// ===== マスタの選択肢 =====

function getPatternOptions() {
  var floors = [];
  var fsh = getSheet_(SHEET_FLOOR);
  var fvals = fsh.getDataRange().getValues();
  for (var i = 1; i < fvals.length; i++) {
    var code = String(fvals[i][0] || '').trim();
    var label = String(fvals[i][2] || '').trim();
    if (code) floors.push({ code: code, label: label || code });
  }
  return { floors: floors };
}

// ===== 勤務パターン =====

function getPatternList() {
  var sh = getSheet_(SHEET_PATTERN);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var values = sh.getRange(2, 1, last - 1, PLAST_COL).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!String(r[PCOL.name - 1] || '').trim()) continue; // 空行は飛ばす
    list.push(rowToPattern_(r, i + 2));
  }
  return list;
}

function rowToPattern_(r, rowIndex) {
  var perStaff = toBool_(r[PCOL.perStaff - 1]);
  return {
    row: rowIndex,
    name: String(r[PCOL.name - 1] || ''),
    start: perStaff ? '' : toTimeString_(r[PCOL.start - 1]),
    end: perStaff ? '' : toTimeString_(r[PCOL.end - 1]),
    floors: String(r[PCOL.floors - 1] || ''),
    desc: String(r[PCOL.desc - 1] || ''),
    overnight: toBool_(r[PCOL.overnight - 1]),
    perStaff: perStaff
  };
}

/**
 * 勤務パターンを保存する。
 * p.row が 0 のときは新規追加、それ以外はその行を更新。
 */
function savePattern(p) {
  requireAdmin_(); // 管理者だけが実行できる
  var sh = getSheet_(SHEET_PATTERN);

  var err = validatePattern_(p, sh);
  if (err) throw new Error(err);

  var row = Number(p.row) || 0;
  if (!row) {
    row = sh.getLastRow() + 1;
    if (row < 2) row = 2;
  }

  var values = [
    String(p.name).trim(),
    p.perStaff ? PER_STAFF_LABEL : fromTimeString_(p.start),
    p.perStaff ? PER_STAFF_LABEL : fromTimeString_(p.end),
    p.floors,
    p.desc,
    !!p.overnight,
    !!p.perStaff
  ];

  sh.getRange(row, 1, 1, PLAST_COL).setValues([values]);
  SpreadsheetApp.flush();
  return getPatternList();
}

/** 入力チェック。問題があればメッセージを返す */
function validatePattern_(p, sh) {
  if (!String(p.name || '').trim()) return 'パターン名を入力してください。';
  if (!String(p.floors || '').trim()) return '対象フロアを選んでください。';

  // パターン名の重複チェック（自分自身は除く）
  var last = sh.getLastRow();
  if (last >= 2) {
    var names = sh.getRange(2, PCOL.name, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      var rowIndex = i + 2;
      if (rowIndex === Number(p.row)) continue;
      if (String(names[i][0] || '').trim() === String(p.name).trim()) {
        return 'パターン名「' + p.name + '」は既に使われています。';
      }
    }
  }

  if (!p.perStaff) {
    var s = String(p.start || '').trim();
    var e = String(p.end || '').trim();
    if (!s || !e) {
      return '開始時刻と終了時刻を入力してください。人によって変わる場合は「勤務時間はスタッフごとに設定」をオンにしてください。';
    }
    if (!isTimeText_(s) || !isTimeText_(e)) {
      return '時刻は 7:30 のような形で入力してください。';
    }
  }

  return null;
}

function isTimeText_(s) {
  return /^([01]?[0-9]|2[0-3]):[0-5][0-9]$/.test(String(s).trim());
}
