/**
 * 質問・要望（フィードバック）の受け付け
 *  ・feedback シートに1件1行で保存する
 *  ・保存のたびに FEEDBACK_MAIL_TO へメールで知らせる
 *  ・管理者は一覧を見て「状態」「返答」を書ける
 */
var FEEDBACK_SHEET = 'feedback';
// 送り先は setting シートの「質問・要望の送り先」に書く（空ならメールは送らず、feedback シートに保存だけする）
var FEEDBACK_MAIL_TO = '';
var FEEDBACK_HEAD = ['日時', 'スタッフNo', '名前', '区分', '画面', '内容', '状態', '返答', '対応日'];
var FEEDBACK_KINDS = ['質問', '不具合', '要望'];
var FEEDBACK_STATUS = ['未対応', '確認中', '対応済み'];

/** feedback シートを返す（無ければ見出しつきで作る） */
function feedbackSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FEEDBACK_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(FEEDBACK_SHEET);
  sh.getRange(1, 1, 1, FEEDBACK_HEAD.length).setValues([FEEDBACK_HEAD]);
  sh.getRange(1, 1, 1, FEEDBACK_HEAD.length).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 130);
  sh.setColumnWidth(6, 400);
  sh.setColumnWidth(8, 300);
  return sh;
}

/** メニュー用：ダイアログで開く */
function openFeedbackDialog() {
  var t = HtmlService.createTemplateFromFile('feedback');
  t.nav = '';
  var html = t.evaluate().setWidth(900).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '質問・要望');
}

/**
 * 画面から呼ぶ：1件保存してメールで知らせる
 * item = { no, name, kind, screen, content }
 */
function submitFeedback(item) {
  item = item || {};
  var kind = String(item.kind || '').trim();
  var content = String(item.content || '').trim();
  var name = String(item.name || '').trim();
  var no = String(item.no || '').trim();
  var screen = String(item.screen || '').trim();
  if (FEEDBACK_KINDS.indexOf(kind) < 0) throw new Error('区分（質問／不具合／要望）を選んでください。');
  if (!name) throw new Error('名前を選んでください。');
  if (!content) throw new Error('内容を入力してください。');
  if (content.length > 2000) throw new Error('内容は2000文字までにしてください。');

  var sh = feedbackSheet_();
  var now = new Date();
  sh.appendRow([now, no, name, kind, screen, content, FEEDBACK_STATUS[0], '', '']);

  var mailOk = true;
  try {
    var when = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    var subject = '【老健シフト】' + kind + '：' + name + (screen ? '（' + screen + '）' : '');
    var body = '老健シフト管理アプリに' + kind + 'が届きました。\n\n'
      + '日時：' + when + '\n'
      + '名前：' + name + (no ? '（' + no + '）' : '') + '\n'
      + '区分：' + kind + '\n'
      + '画面：' + (screen || '（指定なし）') + '\n\n'
      + '----- 内容 -----\n' + content + '\n----------------\n\n'
      + 'スプレッドシートの feedback シート、またはアプリの「質問・要望」画面で確認できます。';
    var mailTo = getSetting_('質問・要望の送り先', FEEDBACK_MAIL_TO);
    if (!mailTo) throw new Error('送り先が未設定');
    MailApp.sendEmail(mailTo, subject, body);
  } catch (e) {
    mailOk = false;   // メールが送れなくても保存は済んでいるので画面には成功で返す
  }
  return JSON.stringify({ ok: true, mail: mailOk });
}

/** 管理者用：一覧を新しい順で返す */
function getFeedbackListJson() {
  requireAdmin_();
  var sh = feedbackSheet_();
  var last = sh.getLastRow();
  if (last < 2) return JSON.stringify([]);
  var v = sh.getRange(2, 1, last - 1, FEEDBACK_HEAD.length).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (!String(v[i][3] || '').trim() && !String(v[i][5] || '').trim()) continue;
    var d = v[i][0];
    out.push({
      row: i + 2,
      when: (d instanceof Date) ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : String(d || ''),
      no: String(v[i][1] || ''),
      name: String(v[i][2] || ''),
      kind: String(v[i][3] || ''),
      screen: String(v[i][4] || ''),
      content: String(v[i][5] || ''),
      status: String(v[i][6] || '') || FEEDBACK_STATUS[0],
      reply: String(v[i][7] || ''),
      done: (v[i][8] instanceof Date) ? Utilities.formatDate(v[i][8], 'Asia/Tokyo', 'yyyy/MM/dd') : String(v[i][8] || '')
    });
  }
  out.reverse();
  return JSON.stringify(out);
}

/** 管理者用：状態と返答を書き込む */
function updateFeedback(row, status, reply) {
  requireAdmin_();
  row = Number(row);
  if (!row || row < 2) throw new Error('行が不正です。');
  status = String(status || '').trim();
  if (FEEDBACK_STATUS.indexOf(status) < 0) throw new Error('状態が不正です。');
  var sh = feedbackSheet_();
  sh.getRange(row, 7).setValue(status);
  sh.getRange(row, 8).setValue(String(reply || '').trim());
  sh.getRange(row, 9).setValue(status === '対応済み' ? new Date() : '');
  return JSON.stringify({ ok: true });
}

/** 送った本人用：自分（スタッフNo）の分だけ新しい順で返す（返答も含む） */
function getMyFeedbackJson(no) {
  no = String(no || '').trim();
  if (!no) return JSON.stringify([]);
  var sh = feedbackSheet_();
  var last = sh.getLastRow();
  if (last < 2) return JSON.stringify([]);
  var v = sh.getRange(2, 1, last - 1, FEEDBACK_HEAD.length).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][1] || '').trim() !== no) continue;
    var d = v[i][0];
    out.push({
      when: (d instanceof Date) ? Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') : String(d || ''),
      kind: String(v[i][3] || ''),
      screen: String(v[i][4] || ''),
      content: String(v[i][5] || ''),
      status: String(v[i][6] || '') || FEEDBACK_STATUS[0],
      reply: String(v[i][7] || ''),
      done: (v[i][8] instanceof Date) ? Utilities.formatDate(v[i][8], 'Asia/Tokyo', 'yyyy/MM/dd') : String(v[i][8] || '')
    });
  }
  out.reverse();
  return JSON.stringify(out);
}
/** 画面用：区分・画面・状態の選択肢 */
function getFeedbackOptionsJson() {
  return JSON.stringify({
    kinds: FEEDBACK_KINDS,
    status: FEEDBACK_STATUS,
    screens: ['スタッフマスタ', '勤務パターン', 'シフト表', 'その日の勤務者', 'シフト希望の申請', 'シフトを自動で組む', 'その他']
  });
}
