function setupStaffSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('staff');
  if (!sh) sh = ss.insertSheet('staff');
  sh.clear();
  const headers = ['スタッフNo','氏名','メールアドレス','担当フロア(主)','担当フロア(兼務)','雇用形態','時給','年収区分','夜勤可否','夜勤専従','夜勤リーダー可否','稼働不可曜日','希望休上限','人員カウント対象外','在籍状況'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  Logger.log('done: ' + sh.getName());
}

function setupPatternSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('pattern');
  if (!sh) sh = ss.insertSheet('pattern');
  sh.clear();
  const headers = ['パターン名','開始時刻','終了時刻','説明'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#e8eaed');
  const rows = [
    ['A','7:30','16:30','早番(通常)'],
    ['A前倒し','7:00','16:00','夜勤4人体制の早番'],
    ['B','11:00','20:00','遅番'],
    ['C','8:30','17:30','5階・新館'],
    ['F','9:00','18:00','フリー業務'],
    ['夜勤','16:00','9:30','日をまたぎます']
    ];
  sh.getRange(2,1,rows.length,4).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,4);
  Logger.log('done: ' + sh.getName());
  }

function setupFloorSheet() {
const ss = SpreadsheetApp.getActiveSpreadsheet();
let sh = ss.getSheetByName('floor');
if (!sh) sh = ss.insertSheet('floor');
sh.clear();
const headers = ['フロア記号介護','フロア記号看護','名称'];
sh.getRange(1,1,1,3).setValues([headers]);
sh.getRange(1,1,1,3).setFontWeight('bold').setBackground('#e8eaed');
const rows = [
['3','3','3階'],
['4','4','4階'],
['5','56','5階'],
['S','南','南館'],
['N','新','新館']
];
sh.getRange(2,1,rows.length,3).setValues(rows);
sh.setFrozenRows(1);
sh.autoResizeColumns(1,3);
Logger.log('done: ' + sh.getName());
}

/** 【検証用】ログインしているユーザーのメールアドレスが取れるか確かめる */
function whoAmI() {
  var active = '';
  var effective = '';
  try { active = Session.getActiveUser().getEmail(); } catch (e) { active = 'エラー: ' + e.message; }
  try { effective = Session.getEffectiveUser().getEmail(); } catch (e) { effective = 'エラー: ' + e.message; }
  var msg = 'getActiveUser  : [' + active + ']\n'
          + 'getEffectiveUser: [' + effective + ']';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}
