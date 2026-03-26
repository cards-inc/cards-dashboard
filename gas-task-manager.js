/**
 * タスク管理ダッシュボード GAS スクリプト
 *
 * 設定手順:
 * 1. スプレッドシートのメニュー → 拡張機能 → Apps Script
 * 2. このスクリプト全体を貼り付けて保存
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（匿名を含む）
 * 4. デプロイURLを index.html の GAS_URL に貼り付ける
 */

// ★ スプレッドシートID（URLの /d/XXXXX/edit の XXXXX 部分）
var SPREADSHEET_ID = '1I8p4IVrMTEKHRdfLwwpW6jVty5N2DIlT7riHC0TW-xw';

// タスクシートのGID（URLの #gid= の値）
var TASK_SHEET_GID       = 235656541;
var EC_EVENTS_SHEET_GID  = 365675538;  // ECイベントカレンダー設定シート
var AUTOMATION_SHEET_GID = 1898911664; // オートメーションシート
var SETTING_SHEET_GID    = 939802887;  // 設定シート（イベント名マスタ）

// オートメーションシート列定義（1始まり）
var AUTO_COL = {
  CATEGORIES:     2,  // B: カテゴリ（カンマ区切り）
  ASSIGNEE:       3,  // C: 担当者
  EC_EVENT:       4,  // D: ECイベント連動
  TASK_NAME:      5,  // E: タスク内容
  DETAIL:         6,  // F: 詳細
  REPEAT:         7,  // G: 繰り返し
  START_DATE:     8,  // H: 開始日
  DUE_DAYS:       9,  // I: タスク期限（日数）
  LAST_GENERATED: 10, // J: 最終生成日（GAS管理）
  NEXT_DUE:       11, // K: 次回生成日（GAS管理）
};

// シート列定義（1始まり）
// A=1(空), B=2(No.), C=3(記載日), D=4(カテゴリ), E=5(ステータス),
// F=6(タスク内容), G=7(詳細), H=8(期限日), I=9(実働時間h), J=10(完了日), K=11(担当者)
var COL = {
  NO:         2,
  ENTRY_DATE: 3,
  CATEGORY:   4,
  STATUS:     5,
  TASK_NAME:  6,
  DETAIL:     7,
  DUE_DATE:   8,
  WORK_HOURS: 9,
  COMPLETED:  10,
  ASSIGNEE:   11,
};

// GIDでシートを取得する（スタンドアロン／コンテナバインド両対応）
function getSheetByGid(gid) {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートを開けません。SPREADSHEET_ID を確認してください');
  var sheets = ss.getSheets();
  var sheet  = sheets.find(function(s) { return s.getSheetId() === Number(gid); });
  if (!sheet) throw new Error('GID ' + gid + ' のシートが見つかりません（シートID一覧: ' + sheets.map(function(s) { return s.getSheetId(); }).join(', ') + '）');
  return sheet;
}

// ──────────────────────────────
// GET: すべてのリクエストをdoGetで処理（JSONP対応）
// ──────────────────────────────
function doGet(e) {
  var action   = e.parameter.action   || '';
  var callback = e.parameter.callback || '';

  var result;
  try {
    if (action === 'add') {
      var data = JSON.parse(e.parameter.data || '{}');
      result = addTask(data);
    } else if (action === 'edit') {
      var data = JSON.parse(e.parameter.data || '{}');
      result = editTask(data);
    } else if (action === 'delete') {
      var data = JSON.parse(e.parameter.data || '{}');
      result = deleteTask(data);
    } else if (action === 'getAutoTemplates') {
      result = { status: 'ok', templates: getAutoTemplates() };
    } else if (action === 'saveAutoTemplates') {
      var data = JSON.parse(e.parameter.data || '{}');
      saveAutoTemplatesData(data.templates || []);
      result = { status: 'ok' };
    } else if (action === 'getSettingEvents') {
      result = { status: 'ok', events: getSettingEvents() };
    } else if (action === 'getEcEvents') {
      result = { status: 'ok', events: getEcEventsFromSheet() };
    } else if (action === 'replaceEcEvents') {
      var data = JSON.parse(e.parameter.data || '{}');
      replaceEcEvents(data);
      result = { status: 'ok' };
    } else if (action === 'postSlack') {
      var data = JSON.parse(e.parameter.data || '{}');
      result = postSlackReport(data);
    } else {
      result = { status: 'error', message: '不明なアクション: ' + action };
    }
  } catch(err) {
    result = { status: 'error', message: err.message };
  }

  return jsonpResponse(result, callback);
}

// ──────────────────────────────
// POST: 互換性のために残す
// ──────────────────────────────
function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch(_) {
    try { data = JSON.parse(e.parameter.data); } catch(__) { data = {}; }
  }

  var result;
  try {
    if (data.action === 'add')       result = addTask(data);
    else if (data.action === 'edit') result = editTask(data);
    else throw new Error('不明なアクション');
  } catch(err) {
    result = { status: 'error', message: err.message };
  }

  return jsonpResponse(result, '');
}

// ──────────────────────────────
// タスク追加
// ──────────────────────────────
function addTask(data) {
  var sheet  = getSheetByGid(TASK_SHEET_GID);
  var nextNo = getLastNo(sheet) + 1;
  var row    = buildRow(nextNo, data);
  sheet.appendRow(row);
  return { status: 'ok', no: nextNo };
}

// ──────────────────────────────
// タスク削除
// ──────────────────────────────
function deleteTask(data) {
  var sheet  = getSheetByGid(TASK_SHEET_GID);
  var taskNo = Number(data.taskNo);
  if (!taskNo) throw new Error('taskNo が指定されていません');

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) throw new Error('No.' + taskNo + ' の行が見つかりません');
  var nos = sheet.getRange(4, COL.NO, lastRow - 3, 1).getValues();
  for (var i = 0; i < nos.length; i++) {
    if (Number(nos[i][0]) === taskNo) {
      sheet.deleteRow(i + 4);
      return { status: 'ok', no: taskNo };
    }
  }
  throw new Error('No.' + taskNo + ' の行が見つかりません');
}

// ──────────────────────────────
// タスク編集
// ──────────────────────────────
function editTask(data) {
  var sheet  = getSheetByGid(TASK_SHEET_GID);
  var taskNo = Number(data.taskNo);
  if (!taskNo) throw new Error('taskNo が指定されていません');

  var lastRow = sheet.getLastRow();
  if (lastRow < 4) throw new Error('No.' + taskNo + ' の行が見つかりません');
  var nos = sheet.getRange(4, COL.NO, lastRow - 3, 1).getValues();
  var targetRow = -1;
  for (var i = 0; i < nos.length; i++) {
    if (Number(nos[i][0]) === taskNo) { targetRow = i + 4; break; }
  }
  if (targetRow === -1) throw new Error('No.' + taskNo + ' の行が見つかりません');

  var row = buildRow(taskNo, data);
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return { status: 'ok', no: taskNo };
}

// ──────────────────────────────
// 行データ生成
// ──────────────────────────────
function buildRow(no, data) {
  return [
    '',                            // A: 空
    no,                            // B: No.
    parseDate(data.entryDate),     // C: 記載日
    data.category    || '',        // D: カテゴリ
    data.status      || '未着手', // E: ステータス
    data.taskName    || '',        // F: タスク内容
    data.detail      || '',        // G: 詳細
    parseDate(data.dueDate),       // H: 期限日
    data.workHours   || '',        // I: 実働時間(h)
    parseDate(data.completedDate), // J: 完了日
    data.assignee    || '',        // K: 担当者
  ];
}

function parseDate(val) {
  if (!val) return '';
  var d = new Date(val);
  if (isNaN(d.getTime())) return '';
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function getLastNo(sheet) {
  var lastRow = sheet.getLastRow();
  for (var r = lastRow; r >= 4; r--) {
    var val = Number(sheet.getRange(r, COL.NO).getValue());
    if (val > 0) return val;
  }
  return 0;
}

// ──────────────────────────────
// レスポンス
// ──────────────────────────────
function jsonpResponse(obj, callback) {
  var json = JSON.stringify(obj);
  var body = callback ? callback + '(' + json + ')' : json;
  var mime = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mime);
}

// ──────────────────────────────
// 設定シートからイベント名マスタを取得（gid=939802887）
// Row1=ヘッダー, Row2+=データ, D列=楽天イベント, E列=Amazonイベント
// ──────────────────────────────
function getSettingEvents() {
  try {
    var sheet = getSheetByGid(SETTING_SHEET_GID);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { rakuten: [], amazon: [] };
    var data = sheet.getRange(2, 4, lastRow - 1, 2).getValues(); // D〜E列
    var rakuten = [], amazon = [];
    data.forEach(function(row) {
      var r = String(row[0] || '').trim();
      var a = String(row[1] || '').trim();
      if (r && rakuten.indexOf(r) === -1) rakuten.push(r);
      if (a && amazon.indexOf(a) === -1) amazon.push(a);
    });
    return { rakuten: rakuten, amazon: amazon };
  } catch(e) {
    Logger.log('getSettingEvents error: ' + e);
    return { rakuten: [], amazon: [] };
  }
}

// ──────────────────────────────
// ECイベントシート管理（gid=365675538）
// 構造: Row1=セクションラベル(楽天市場/Amazon), Row2=列ヘッダ, Row3+=データ
// 列: A(空), B=楽天イベント名, C=楽天開始日, D=楽天終了日, E(空), F=Amazonイベント名, G=Amazon開始日, H=Amazon終了日
// ──────────────────────────────
function getEcEventsFromSheet() {
  try {
    var sheet = getSheetByGid(EC_EVENTS_SHEET_GID);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    var data = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
    var events = [];
    var fmtD = function(v) {
      if (!v) return null;
      try { return Utilities.formatDate(new Date(v), 'Asia/Tokyo', 'yyyy-MM-dd'); } catch(_) { return null; }
    };
    data.forEach(function(row) {
      if (row[1]) events.push({ label: String(row[1]), start: fmtD(row[2]), end: fmtD(row[3]), type: 'rakuten' });
      if (row[5]) events.push({ label: String(row[5]), start: fmtD(row[6]), end: fmtD(row[7]), type: 'amazon' });
    });
    return events;
  } catch(e) { Logger.log('getEcEventsFromSheet error: ' + e); return []; }
}

function replaceEcEvents(data) {
  var sheet = getSheetByGid(EC_EVENTS_SHEET_GID);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 3) sheet.getRange(3, 1, lastRow - 2, 8).clearContent();
  var events = data.events || [];
  var rakuten = events.filter(function(e) { return e.type === 'rakuten'; });
  var amazon  = events.filter(function(e) { return e.type === 'amazon'; });
  var len = Math.max(rakuten.length, amazon.length);
  if (!len) return { status: 'ok' };
  var rows = [];
  for (var i = 0; i < len; i++) {
    var r = rakuten[i] || {}, a = amazon[i] || {};
    rows.push(['', r.label||'', r.start||'', r.end||'', '', a.label||'', a.start||'', a.end||'']);
  }
  sheet.getRange(3, 1, rows.length, 8).setValues(rows);
  return { status: 'ok' };
}

// ──────────────────────────────
// オートメーションシート管理（gid=1898911664）
// Row1=タイトル, Row2=ヘッダ, Row3+=データ
// ──────────────────────────────
function getAutoTemplates() {
  try {
    var sheet = getSheetByGid(AUTOMATION_SHEET_GID);
    var lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    var data = sheet.getRange(3, 1, lastRow - 2, 11).getValues();
    return data
      .map(function(row, i) {
        if (!row[AUTO_COL.TASK_NAME - 1]) return null;
        var catsRaw = String(row[AUTO_COL.CATEGORIES - 1] || '');
        var startVal    = row[AUTO_COL.START_DATE - 1];
        var lastGenVal  = row[AUTO_COL.LAST_GENERATED - 1];
        var nextDueVal  = row[AUTO_COL.NEXT_DUE - 1];
        return {
          _row:          i + 3,
          categories:    catsRaw ? catsRaw.split(',').map(function(s){ return s.trim(); }) : [],
          assignee:      String(row[AUTO_COL.ASSIGNEE - 1] || ''),
          ecEventLabel:  row[AUTO_COL.EC_EVENT - 1] || null,
          taskName:      String(row[AUTO_COL.TASK_NAME - 1] || ''),
          detail:        String(row[AUTO_COL.DETAIL - 1] || ''),
          repeat:        String(row[AUTO_COL.REPEAT - 1] || 'weekly'),
          startDate:     startVal    ? new Date(startVal).toISOString()    : null,
          dueDays:       parseInt(row[AUTO_COL.DUE_DAYS - 1]) || 7,
          lastGenerated: lastGenVal  ? new Date(lastGenVal).toISOString()  : null,
          nextDue:       nextDueVal  ? new Date(nextDueVal).toISOString()  : null,
        };
      })
      .filter(Boolean);
  } catch(e) { Logger.log('getAutoTemplates error: ' + e); return []; }
}

function saveAutoTemplatesData(templates) {
  try {
    var sheet = getSheetByGid(AUTOMATION_SHEET_GID);
    var lastRow = sheet.getLastRow();
    if (lastRow > 2) sheet.getRange(3, 1, lastRow - 2, 11).clearContent();
    if (!templates.length) return;
    var rows = templates.map(function(tmpl) {
      var cats = Array.isArray(tmpl.categories) ? tmpl.categories : (tmpl.category ? [tmpl.category] : []);
      return [
        '',
        cats.join(', '),
        tmpl.assignee      || '',
        tmpl.ecEventLabel  || '',
        tmpl.taskName      || '',
        tmpl.detail        || '',
        tmpl.repeat        || 'weekly',
        tmpl.startDate     ? new Date(tmpl.startDate)     : '',
        tmpl.dueDays       || 7,
        tmpl.lastGenerated ? new Date(tmpl.lastGenerated) : '',
        tmpl.nextDue       ? new Date(tmpl.nextDue)       : '',
      ];
    });
    sheet.getRange(3, 1, rows.length, 11).setValues(rows);
  } catch(e) { Logger.log('saveAutoTemplatesData error: ' + e); }
}

// 毎日9時に実行: nextDue <= 今日 のテンプレートからタスクを自動生成
function runAutoTemplates() {
  var today = new Date(); today.setHours(0,0,0,0);
  var templates = getAutoTemplates();
  var ecEvents  = getEcEventsFromSheet();
  var autoSheet = getSheetByGid(AUTOMATION_SHEET_GID);
  var taskSheet = getSheetByGid(TASK_SHEET_GID);

  templates.forEach(function(tmpl) {
    if (!tmpl.nextDue) return;
    var nextDue = new Date(tmpl.nextDue); nextDue.setHours(0,0,0,0);

    // ECイベント連動: イベント開始の7営業日前にタスク生成
    // 通常繰り返し: nextDue当日に生成
    var generateDate = tmpl.ecEventLabel
      ? subtractBusinessDaysGas(new Date(nextDue), 7)
      : new Date(nextDue);
    if (generateDate > today) return;

    var lastGen = tmpl.lastGenerated ? new Date(tmpl.lastGenerated) : null;
    if (lastGen) { lastGen.setHours(0,0,0,0); }
    if (lastGen && lastGen >= generateDate) return;

    var cats = Array.isArray(tmpl.categories) ? tmpl.categories : [];
    var dueDate = tmpl.ecEventLabel
      ? subtractBusinessDaysGas(new Date(nextDue), 2)
      : (function(){ var d = new Date(nextDue); d.setDate(d.getDate() + (tmpl.dueDays||7)); return d; })();
    var entryDate = tmpl.ecEventLabel
      ? generateDate
      : new Date(nextDue);

    cats.forEach(function(cat) {
      var nextNo = getLastNo(taskSheet) + 1;
      taskSheet.appendRow(buildRow(nextNo, {
        entryDate:     Utilities.formatDate(entryDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
        category:      cat,
        status:        '未着手',
        taskName:      tmpl.taskName,
        detail:        tmpl.detail || '',
        dueDate:       Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
        assignee:      tmpl.assignee || '',
        workHours:     0,
        completedDate: '',
      }));
    });

    autoSheet.getRange(tmpl._row, AUTO_COL.LAST_GENERATED).setValue(generateDate);

    var newNextDue = tmpl.ecEventLabel
      ? getNextEcEventDateFromSheet(tmpl.ecEventLabel, nextDue, ecEvents)
      : computeNextRepeatDate(tmpl.startDate, tmpl.repeat, nextDue);
    if (newNextDue) autoSheet.getRange(tmpl._row, AUTO_COL.NEXT_DUE).setValue(newNextDue);
  });
}

function getNextEcEventDateFromSheet(ecEventLabel, afterDate, ecEvents) {
  if (!ecEvents || !ecEvents.length) return null;
  var after = new Date(afterDate); after.setHours(0,0,0,0);
  var baseLabel = ecEventLabel.replace('×5,0日','');
  var futures = ecEvents
    .filter(function(e) { return e.start && (e.label === ecEventLabel || e.label === baseLabel); })
    .map(function(e) { var d = new Date(e.start); d.setHours(0,0,0,0); return d; })
    .filter(function(d) { return d > after; })
    .sort(function(a,b) { return a - b; });
  return futures[0] || null;
}

function computeNextRepeatDate(startDate, repeat, afterDate) {
  var ref = new Date(afterDate); ref.setHours(0,0,0,0);
  var d = new Date(startDate); d.setHours(0,0,0,0);
  while (d <= ref) {
    if (repeat === 'daily')         d.setDate(d.getDate() + 1);
    else if (repeat === 'weekly')   d.setDate(d.getDate() + 7);
    else if (repeat === 'biweekly') d.setDate(d.getDate() + 14);
    else if (repeat === 'monthly')  d.setMonth(d.getMonth() + 1);
    else break;
  }
  return d > ref ? d : null;
}

function subtractBusinessDaysGas(date, days) {
  var d = new Date(date);
  var count = 0;
  while (count < days) {
    d.setDate(d.getDate() - 1);
    var dow = d.getDay();
    if (dow !== 0 && dow !== 6) count++;
  }
  return d;
}

// ──────────────────────────────
// Slack 業務報告投稿
// ──────────────────────────────
// GASスクリプトプロパティに SLACK_BOT_TOKEN を設定すること
var SLACK_CHANNEL_ID = 'C071F406H5H';
var SLACK_REPORT_MENTION = '<!subteam^S0ANAB1BZA7>'; // @report ユーザーグループ

function postSlackReport(data) {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) throw new Error('SLACK_BOT_TOKEN がスクリプトプロパティに設定されていません');

  var reportText = data.text || '';
  if (!reportText.trim()) throw new Error('投稿テキストが空です');

  var reportDate = data.date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  // チャンネル履歴から当日の業務報告スレッドを検索
  var threadTs = findReportThread(token, reportDate);
  if (!threadTs) throw new Error('本日の業務報告スレッド（業務報告bot投稿）が見つかりません。12時以降に再試行してください');

  // @report メンション付きでスレッドに返信
  var fullText = SLACK_REPORT_MENTION + '\n' + reportText;

  var postRes = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' },
    payload: JSON.stringify({
      channel: SLACK_CHANNEL_ID,
      thread_ts: threadTs,
      text: fullText,
    }),
    muteHttpExceptions: true,
  });
  var postJson = JSON.parse(postRes.getContentText());
  if (!postJson.ok) throw new Error('Slack投稿エラー: ' + (postJson.error || '不明'));

  return { status: 'ok', ts: postJson.ts, thread_ts: threadTs };
}

function findReportThread(token, dateStr) {
  // dateStr = 'yyyy-MM-dd' → 当日 00:00〜23:59 (JST = UTC+9)
  var d = new Date(dateStr + 'T00:00:00+09:00');
  var oldest = Math.floor(d.getTime() / 1000);
  var latest = oldest + 86400;

  var url = 'https://slack.com/api/conversations.history'
    + '?channel=' + SLACK_CHANNEL_ID
    + '&oldest=' + oldest
    + '&latest=' + latest
    + '&limit=50';

  var res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  var json = JSON.parse(res.getContentText());
  if (!json.ok) {
    Logger.log('conversations.history error: ' + json.error);
    return null;
  }

  // 業務報告botの投稿（B0A0C056AMU）から「業務報告」を含むメッセージのtsを返す
  var msgs = json.messages || [];
  for (var i = 0; i < msgs.length; i++) {
    var m = msgs[i];
    if (m.text && m.text.indexOf('業務報告') >= 0 && m.bot_id === 'B0A0C056AMU') {
      return m.ts;
    }
  }
  return null;
}

// ──────────────────────────────
// 期限切れタスク通知（Slack #check_alert）
// ──────────────────────────────
var ALERT_CHANNEL_ID = 'C09UU758NS3';

// 担当者名 → SlackユーザーID マッピング
var ASSIGNEE_SLACK_MAP = {
  '佐藤敦子': 'U09TA7UDED7',
  '松下絵里': 'U09SV9QHWCE',
  '樽井翔平': 'U0AC5TSTMFH',
  '木口亜理沙': 'U09PYRF9WBW',
  '小池通子': 'U0A713KLGKV',
  '中瀬英輔': 'U0ABSAF8KJ7',
  '堀田希': 'U0AH88CEWAE',
  '鈴木詩乃': 'U0AHPA8KPE3',
  '秋葉唯': 'U0711EST0NA',
  '増田真喜': 'U0A5EE1NA3D',
  '佐渡明日香': 'U0934M42QCX',
  '上原起': 'U08BB4MH8P3',
};

function notifyOverdueTasks() {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) { Logger.log('SLACK_BOT_TOKEN が未設定'); return; }

  var sheet = getSheetByGid(TASK_SHEET_GID);
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  var data = sheet.getRange(4, 1, lastRow - 3, 11).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);

  // 担当者ごとに期限切れタスクをグループ化
  var overdueByPerson = {};
  data.forEach(function(row) {
    var status  = String(row[COL.STATUS - 1] || '');
    var dueDate = row[COL.DUE_DATE - 1];
    if (status === '対応完了' || !dueDate) return;
    var due = new Date(dueDate); due.setHours(0, 0, 0, 0);
    if (due >= today) return;

    var assignee = String(row[COL.ASSIGNEE - 1] || '').trim();
    var taskName = String(row[COL.TASK_NAME - 1] || '');
    var dueFmt   = Utilities.formatDate(due, 'Asia/Tokyo', 'M/d');
    if (!assignee) assignee = '未割当';

    if (!overdueByPerson[assignee]) overdueByPerson[assignee] = [];
    overdueByPerson[assignee].push({ name: taskName, due: dueFmt });
  });

  var people = Object.keys(overdueByPerson);
  if (!people.length) {
    Logger.log('期限切れタスクなし');
    return;
  }

  // メッセージ組み立て
  var totalCount = 0;
  var lines = [];
  people.sort().forEach(function(person) {
    var tasks = overdueByPerson[person];
    totalCount += tasks.length;
    var slackId = ASSIGNEE_SLACK_MAP[person.replace(/\s/g, '')];
    var mention = slackId ? '<@' + slackId + '>' : person;
    lines.push('');
    lines.push(mention + '（' + tasks.length + '件）');
    tasks.forEach(function(t) {
      lines.push('　・' + t.name + '（期限: ' + t.due + '）');
    });
  });

  var text = ':warning: *期限切れタスクが ' + totalCount + ' 件あります*' + lines.join('\n');

  UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' },
    payload: JSON.stringify({ channel: ALERT_CHANNEL_ID, text: text }),
    muteHttpExceptions: true,
  });

  Logger.log('期限切れ通知送信: ' + totalCount + '件');
}

// トリガー設定（GASエディタから1度だけ手動実行）
function setupTriggers() {
  // 既存トリガーをクリア
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'runAutoTemplates' || fn === 'notifyOverdueTasks') ScriptApp.deleteTrigger(t);
  });
  // オートメーション: 毎朝9時
  ScriptApp.newTrigger('runAutoTemplates')
    .timeBased().atHour(9).everyDays(1).inTimezone('Asia/Tokyo').create();
  // 期限切れ通知: 毎朝9時
  ScriptApp.newTrigger('notifyOverdueTasks')
    .timeBased().atHour(9).everyDays(1).inTimezone('Asia/Tokyo').create();
  Logger.log('トリガー設定完了（オートメーション9時 / 期限切れ通知9時）');
}
