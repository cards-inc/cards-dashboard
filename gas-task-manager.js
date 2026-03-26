/**
 * タスク管理ダッシュボード GAS スクリプト
 *
 * 設定手順:
 * 1. スプレッドシートのメニュー → 拡張機能 → Apps Script
 * 2. このスクリプト全体を貼り付けて保存
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（匿名を含む）
 * 4. デプロイURLを task-dashboard.html の GAS_URL に貼り付ける
 */

// ★ スプレッドシートID（URLの /d/XXXXX/edit の XXXXX 部分）
const SPREADSHEET_ID = '1I8p4IVrMTEKHRdfLwwpW6jVty5N2DIlT7riHC0TW-xw';

// タスクシートのGID（URLの #gid= の値）
const TASK_SHEET_GID       = 235656541;
const EC_EVENTS_SHEET_GID  = 365675538;  // ECイベントカレンダー設定シート
const AUTOMATION_SHEET_GID = 1898911664; // オートメーションシート

// オートメーションシート列定義（1始まり）
const AUTO_COL = {
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
const COL = {
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
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートを開けません。SPREADSHEET_ID を確認してください');
  const sheets = ss.getSheets();
  const sheet  = sheets.find(s => s.getSheetId() === Number(gid));
  if (!sheet) throw new Error('GID ' + gid + ' のシートが見つかりません（シートID一覧: ' + sheets.map(s => s.getSheetId()).join(', ') + '）');
  return sheet;
}

// ──────────────────────────────
// GET: すべてのリクエストをdoGetで処理（JSONP対応）
// file://からでも確実に動作する
// ──────────────────────────────
function doGet(e) {
  const action   = e.parameter.action   || '';
  const callback = e.parameter.callback || '';

  let result;
  try {
    if (action === 'add') {
      const data = JSON.parse(e.parameter.data || '{}');
      result = addTask(data);
    } else if (action === 'edit') {
      const data = JSON.parse(e.parameter.data || '{}');
      result = editTask(data);
    } else if (action === 'delete') {
      const data = JSON.parse(e.parameter.data || '{}');
      result = deleteTask(data);
    } else if (action === 'postSlack') {
      const data = JSON.parse(e.parameter.data || '{}');
      result = postToSlack(data.text, data.thread_ts);
    } else if (action === 'findThread') {
      return findSlackThread(e);
    } else if (action === 'scanMinutes') {
      result = scanMeetingMinutes();
    } else if (action === 'getAutoTemplates') {
      result = { status: 'ok', templates: getAutoTemplates() };
    } else if (action === 'saveAutoTemplates') {
      const data = JSON.parse(e.parameter.data || '{}');
      saveAutoTemplatesData(data.templates || []);
      result = { status: 'ok' };
    } else if (action === 'getEcEvents') {
      result = { status: 'ok', events: getEcEventsFromSheet() };
    } else if (action === 'replaceEcEvents') {
      const data = JSON.parse(e.parameter.data || '{}');
      replaceEcEvents(data);
      result = { status: 'ok' };
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
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch(_) {
    try { data = JSON.parse(e.parameter.data); } catch(__) { data = {}; }
  }

  let result;
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
  const sheet  = getSheetByGid(TASK_SHEET_GID);
  const nextNo = getLastNo(sheet) + 1;
  const row    = buildRow(nextNo, data);
  sheet.appendRow(row);
  return { status: 'ok', no: nextNo };
}

// ──────────────────────────────
// タスク削除
// ──────────────────────────────
function deleteTask(data) {
  const sheet  = getSheetByGid(TASK_SHEET_GID);
  const taskNo = Number(data.taskNo);
  if (!taskNo) throw new Error('taskNo が指定されていません');

  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    if (Number(sheet.getRange(r, COL.NO).getValue()) === taskNo) {
      sheet.deleteRow(r);
      return { status: 'ok', no: taskNo };
    }
  }
  throw new Error('No.' + taskNo + ' の行が見つかりません');
}

// ──────────────────────────────
// タスク編集
// ──────────────────────────────
function editTask(data) {
  const sheet  = getSheetByGid(TASK_SHEET_GID);
  const taskNo = Number(data.taskNo);
  if (!taskNo) throw new Error('taskNo が指定されていません');

  const lastRow = sheet.getLastRow();
  let targetRow = -1;
  for (let r = 2; r <= lastRow; r++) {
    if (Number(sheet.getRange(r, COL.NO).getValue()) === taskNo) {
      targetRow = r;
      break;
    }
  }
  if (targetRow === -1) throw new Error('No.' + taskNo + ' の行が見つかりません');

  const row = buildRow(taskNo, data);
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
  const d = new Date(val);
  if (isNaN(d.getTime())) return '';
  // 時間を除いた日付のみ返す
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function getLastNo(sheet) {
  const lastRow = sheet.getLastRow();
  for (let r = lastRow; r >= 2; r--) {
    const val = Number(sheet.getRange(r, COL.NO).getValue());
    if (val > 0) return val;
  }
  return 0;
}

// ──────────────────────────────
// レスポンス
// ──────────────────────────────
function jsonpResponse(obj, callback) {
  const json = JSON.stringify(obj);
  const body = callback ? callback + '(' + json + ')' : json;
  const mime = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(body).setMimeType(mime);
}

// ──────────────────────────────
// Slack連携
// ──────────────────────────────
const SLACK_WEBHOOK_URL = ''; // ← Slack Incoming Webhook URL を設定
const SLACK_BOT_TOKEN   = ''; // ← Slack Bot Token (xoxb-...) を設定
const SLACK_CHANNEL_ID  = ''; // ← 投稿先チャンネルID を設定
// 業務報告投稿時に先頭へ付けるメンション（例: '@here', '@channel', '@username'）
// 不要な場合は空文字にする
const SLACK_MENTION     = '@report';

function postToSlack(text, thread_ts) {
  if (!SLACK_WEBHOOK_URL) throw new Error('SLACK_WEBHOOK_URL が未設定です');
  const withMention = SLACK_MENTION ? SLACK_MENTION + '\n' + text : text;
  const resolved = resolveSlackMentions(withMention);
  const payload = { text: resolved };
  if (thread_ts) payload.thread_ts = thread_ts;
  const res = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
  });
  if (res.getResponseCode() !== 200) throw new Error('Slack投稿失敗: ' + res.getContentText());
  return { status: 'ok' };
}

// @handle を Slack メンション形式に変換
function resolveSlackMentions(text) {
  return text.replace(/@(\w+)/g, (match, handle) => {
    if (handle === 'here')     return '<!here>';
    if (handle === 'channel')  return '<!channel>';
    if (handle === 'everyone') return '<!everyone>';
    // ユーザーグループを先に検索
    const groupId = lookupSlackHandle(handle);
    if (groupId) return `<!subteam^${groupId}|@${handle}>`;
    // 個人ユーザーを検索
    const userId = lookupSlackUser(handle);
    if (userId) return `<@${userId}>`;
    return match;
  });
}

// Slack ユーザーグループのハンドル → ID を検索
const _handleCache = {};
function lookupSlackHandle(handle) {
  if (_handleCache[handle] !== undefined) return _handleCache[handle];
  try {
    const res  = UrlFetchApp.fetch(
      'https://slack.com/api/usergroups.list?include_disabled=false',
      { headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN } }
    );
    const json = JSON.parse(res.getContentText());
    if (json.ok && json.usergroups) {
      json.usergroups.forEach(g => { _handleCache[g.handle] = g.id; });
    }
  } catch(_) {}
  return _handleCache[handle] || null;
}

// Slack 個人ユーザーのハンドル（display_name / name） → ID を検索
const _userCache = {};
function lookupSlackUser(handle) {
  if (_userCache[handle] !== undefined) return _userCache[handle];
  try {
    const res  = UrlFetchApp.fetch(
      'https://slack.com/api/users.list?limit=200',
      { headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN } }
    );
    const json = JSON.parse(res.getContentText());
    if (json.ok && json.members) {
      json.members.forEach(m => {
        const dn = (m.profile && m.profile.display_name) ? m.profile.display_name : '';
        if (m.name)  _userCache[m.name]  = m.id;
        if (dn)      _userCache[dn]      = m.id;
      });
    }
  } catch(_) {}
  return _userCache[handle] || null;
}

function findSlackThread(e) {
  const callback = e.parameter.callback || '';
  if (!SLACK_BOT_TOKEN || !SLACK_CHANNEL_ID) {
    return jsonpResponse({ status: 'error', message: 'Slack設定が未完了です' }, callback);
  }
  const res  = UrlFetchApp.fetch(
    'https://slack.com/api/conversations.history?channel=' + SLACK_CHANNEL_ID + '&limit=100',
    { headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN } }
  );
  const json = JSON.parse(res.getContentText());
  if (!json.ok) {
    return jsonpResponse({ status: 'error', message: 'Slack APIエラー: ' + json.error }, callback);
  }
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');
  const messages = json.messages || [];
  const msg = messages.find(m =>
    m.text && m.text.includes('業務報告') && m.text.includes(today)
  );
  if (!msg) {
    const sample = messages.slice(0, 3).map(m => (m.text || '').substring(0, 60));
    return jsonpResponse({
      status: 'ok', thread_ts: null,
      debug: { today: today, count: messages.length, sample: sample }
    }, callback);
  }
  return jsonpResponse({ status: 'ok', thread_ts: msg.ts }, callback);
}

// ──────────────────────────────
// AI議事録からタスク自動追加
// ──────────────────────────────
// 設定: AI議事録フォルダのGoogle Drive フォルダID（DriveのURLの /folders/XXXXX 部分）
const MINUTES_FOLDER_ID = '';   // ← ここにフォルダIDを設定
// 設定: 自分の担当者名（空の場合は全タスクを取得）
const MINUTES_ASSIGNEE  = '';   // ← 例: '山田太郎'
// 処理済みファイルにつけるプレフィックス
const PROCESSED_PREFIX  = '[処理済] ';
// 直近何時間以内のファイルを対象にするか
const MINUTES_HOURS     = 24;

function scanMeetingMinutes() {
  if (!MINUTES_FOLDER_ID) return { status: 'error', message: 'MINUTES_FOLDER_ID が未設定です。GASの定数を設定してください。' };

  const folder  = DriveApp.getFolderById(MINUTES_FOLDER_ID);
  const files   = folder.getFiles();
  const cutoff  = new Date(Date.now() - MINUTES_HOURS * 60 * 60 * 1000);
  const added   = [];
  const skipped = [];

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();

    // 処理済みスキップ
    if (name.startsWith(PROCESSED_PREFIX)) { skipped.push(name); continue; }

    // 古いファイルスキップ
    if (file.getLastUpdated() < cutoff) { skipped.push(name); continue; }

    let content = '';
    try {
      const mime = file.getMimeType();
      if (mime === MimeType.GOOGLE_DOCS) {
        content = DocumentApp.openById(file.getId()).getBody().getText();
      } else if (mime === MimeType.PLAIN_TEXT) {
        content = file.getBlob().getDataAsString('UTF-8');
      } else {
        skipped.push(name + ' (非対応形式)');
        continue;
      }
    } catch(e) {
      skipped.push(name + ' (読み込みエラー: ' + e.message + ')');
      continue;
    }

    // タスク行を抽出
    const taskPatterns = [
      /^[→\-\•□●▶▷]\s*(.+)/,
      /^アクション[:：]\s*(.+)/i,
      /^TODO[:：]\s*(.+)/i,
      /^タスク[:：]\s*(.+)/i,
      /^\[?\s*(?:TODO|todo|task)\s*\]?\s*(.+)/,
    ];

    const lines = content.split('\n');
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.length < 3) continue;

      // 担当者フィルタ（設定時のみ）
      if (MINUTES_ASSIGNEE && !trimmed.includes(MINUTES_ASSIGNEE)) continue;

      let taskName = null;
      for (const pat of taskPatterns) {
        const m = trimmed.match(pat);
        if (m) { taskName = m[1].trim(); break; }
      }
      if (!taskName || taskName.length < 3) continue;

      try {
        const result = addTask({
          entryDate: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
          category:  '会議',
          status:    '未着手',
          taskName:  taskName,
          detail:    '出典: ' + name,
          assignee:  MINUTES_ASSIGNEE || '',
        });
        added.push({ no: result.no, taskName });
      } catch(e) {
        skipped.push('追加失敗: ' + taskName);
      }
    }

    // 処理済みにリネーム
    try { file.setName(PROCESSED_PREFIX + name); } catch(_) {}
  }

  return { status: 'ok', added, skipped };
}

// ──────────────────────────────
// 毎朝Slack通知: 期限切れ・本日期限タスク
// ──────────────────────────────
function sendDailyAlert() {
  const sheet = getSheetByGid(TASK_SHEET_GID);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const lastRow = sheet.getLastRow();
  const overdue = [];
  const dueToday = [];

  for (let r = 2; r <= lastRow; r++) {
    const status   = String(sheet.getRange(r, COL.STATUS).getValue()).trim();
    const taskName = String(sheet.getRange(r, COL.TASK_NAME).getValue()).trim();
    const assignee = String(sheet.getRange(r, COL.ASSIGNEE).getValue()).trim();
    const dueDateVal = sheet.getRange(r, COL.DUE_DATE).getValue();

    if (!taskName || status === '対応完了') continue;
    if (!dueDateVal) continue;

    const due = new Date(dueDateVal);
    due.setHours(0, 0, 0, 0);

    if (due < today) {
      const days = Math.floor((today - due) / 86400000);
      overdue.push({ taskName, assignee, days });
    } else if (due.getTime() === today.getTime()) {
      dueToday.push({ taskName, assignee });
    }
  }

  if (!overdue.length && !dueToday.length) return; // 該当なし → 投稿しない

  const dateLabel = Utilities.formatDate(today, 'Asia/Tokyo', 'M月d日');
  let text = `📋 *タスクアラート｜${dateLabel}*\n`;

  if (dueToday.length) {
    text += `\n📅 *本日期限* (${dueToday.length}件)\n`;
    dueToday.forEach(t => {
      text += `• ${t.taskName}　_${t.assignee}_\n`;
    });
  }

  if (overdue.length) {
    text += `\n⚠️ *期限切れ* (${overdue.length}件)\n`;
    overdue.forEach(t => {
      text += `• ${t.taskName}　_${t.assignee}_ （${t.days}日超過）\n`;
    });
  }

  // 業務報告スレッドを検索して返信、見つからなければスタンドアロンで投稿
  let thread_ts = null;
  try {
    if (SLACK_BOT_TOKEN && SLACK_CHANNEL_ID) {
      const res  = UrlFetchApp.fetch(
        'https://slack.com/api/conversations.history?channel=' + SLACK_CHANNEL_ID + '&limit=100',
        { headers: { Authorization: 'Bearer ' + SLACK_BOT_TOKEN } }
      );
      const json = JSON.parse(res.getContentText());
      if (json.ok) {
        const todayLabel = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年M月d日');
        const msg = (json.messages || []).find(function(m) {
          return m.text && m.text.includes('業務報告') && m.text.includes(todayLabel);
        });
        if (msg) thread_ts = msg.ts;
      }
    }
  } catch(e) { Logger.log('業務報告スレッド検索失敗: ' + e); }

  postToSlack(text, thread_ts);
}

// トリガー設定: 毎日16時に sendDailyAlert を実行（業務報告スレッド返信用）
function setupDailyAlertTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'sendDailyAlert'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendDailyAlert')
    .timeBased().atHour(16).everyDays(1).inTimezone('Asia/Tokyo').create();
  Logger.log('タスクアラートトリガーを設定しました（毎日16時）');
}

// ──────────────────────────────
// ECイベントシート管理（gid=365675538）
// 構造: Row1=セクションラベル(楽天市場/Amazon), Row2=列ヘッダ, Row3+=データ
// 列: A(空), B=楽天イベント名, C=楽天開始日, D=楽天終了日, E(空), F=AmazonIベント名, G=Amazon開始日, H=Amazon終了日
// ──────────────────────────────
function getEcEventsFromSheet() {
  try {
    const sheet = getSheetByGid(EC_EVENTS_SHEET_GID);
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return [];
    const data = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
    const events = [];
    const fmtD = function(v) {
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
  const sheet = getSheetByGid(EC_EVENTS_SHEET_GID);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) sheet.getRange(3, 1, lastRow - 2, 8).clearContent();
  const events = data.events || [];
  const rakuten = events.filter(function(e) { return e.type === 'rakuten'; });
  const amazon  = events.filter(function(e) { return e.type === 'amazon'; });
  const len = Math.max(rakuten.length, amazon.length);
  if (!len) return { status: 'ok' };
  const rows = [];
  for (let i = 0; i < len; i++) {
    const r = rakuten[i] || {}, a = amazon[i] || {};
    rows.push(['', r.label||'', r.start||'', r.end||'', '', a.label||'', a.start||'', a.end||'']);
  }
  sheet.getRange(3, 1, rows.length, 8).setValues(rows);
  return { status: 'ok' };
}

// ──────────────────────────────
// オートメーションシート管理（gid=1898911664）
// Row1=ヘッダ, Row2+=データ
// ──────────────────────────────
function getAutoTemplates() {
  try {
    const sheet = getSheetByGid(AUTOMATION_SHEET_GID);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    return data
      .map(function(row, i) {
        if (!row[AUTO_COL.TASK_NAME - 1]) return null;
        const catsRaw = String(row[AUTO_COL.CATEGORIES - 1] || '');
        const startVal = row[AUTO_COL.START_DATE - 1];
        const lastGenVal = row[AUTO_COL.LAST_GENERATED - 1];
        const nextDueVal = row[AUTO_COL.NEXT_DUE - 1];
        return {
          _row:          i + 2,
          categories:    catsRaw ? catsRaw.split(',').map(function(s){ return s.trim(); }) : [],
          assignee:      String(row[AUTO_COL.ASSIGNEE - 1] || ''),
          ecEventLabel:  row[AUTO_COL.EC_EVENT - 1] || null,
          taskName:      String(row[AUTO_COL.TASK_NAME - 1] || ''),
          detail:        String(row[AUTO_COL.DETAIL - 1] || ''),
          repeat:        String(row[AUTO_COL.REPEAT - 1] || 'weekly'),
          startDate:     startVal ? new Date(startVal).toISOString() : null,
          dueDays:       parseInt(row[AUTO_COL.DUE_DAYS - 1]) || 7,
          lastGenerated: lastGenVal ? new Date(lastGenVal).toISOString() : null,
          nextDue:       nextDueVal ? new Date(nextDueVal).toISOString() : null,
        };
      })
      .filter(Boolean);
  } catch(e) { Logger.log('getAutoTemplates error: ' + e); return []; }
}

function saveAutoTemplatesData(templates) {
  try {
    const sheet = getSheetByGid(AUTOMATION_SHEET_GID);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 11).clearContent();
    if (!templates.length) return;
    const rows = templates.map(function(tmpl) {
      const cats = Array.isArray(tmpl.categories) ? tmpl.categories : (tmpl.category ? [tmpl.category] : []);
      return [
        '',
        cats.join(', '),
        tmpl.assignee || '',
        tmpl.ecEventLabel || '',
        tmpl.taskName || '',
        tmpl.detail || '',
        tmpl.repeat || 'weekly',
        tmpl.startDate ? new Date(tmpl.startDate) : '',
        tmpl.dueDays || 7,
        tmpl.lastGenerated ? new Date(tmpl.lastGenerated) : '',
        tmpl.nextDue ? new Date(tmpl.nextDue) : '',
      ];
    });
    sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  } catch(e) { Logger.log('saveAutoTemplatesData error: ' + e); }
}

// 毎日9時に実行: nextDue <= 今日 のテンプレートからタスクを自動生成
function runAutoTemplates() {
  const today = new Date(); today.setHours(0,0,0,0);
  const templates = getAutoTemplates();
  const ecEvents  = getEcEventsFromSheet();
  const autoSheet = getSheetByGid(AUTOMATION_SHEET_GID);
  const taskSheet = getSheetByGid(TASK_SHEET_GID);

  templates.forEach(function(tmpl) {
    if (!tmpl.nextDue) return;
    const nextDue = new Date(tmpl.nextDue); nextDue.setHours(0,0,0,0);
    if (nextDue > today) return;

    const lastGen = tmpl.lastGenerated ? new Date(tmpl.lastGenerated) : null;
    if (lastGen) { lastGen.setHours(0,0,0,0); }
    if (lastGen && lastGen >= nextDue) return; // 生成済み

    const cats = Array.isArray(tmpl.categories) ? tmpl.categories : [];
    const dueDate = tmpl.ecEventLabel
      ? subtractBusinessDaysGas(new Date(nextDue), 2)
      : (function(){ var d = new Date(nextDue); d.setDate(d.getDate() + (tmpl.dueDays||7)); return d; })();
    const entryDate = tmpl.ecEventLabel
      ? subtractBusinessDaysGas(new Date(nextDue), 5)
      : new Date(nextDue);

    cats.forEach(function(cat) {
      const nextNo = getLastNo(taskSheet) + 1;
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

    // lastGenerated を更新
    autoSheet.getRange(tmpl._row, AUTO_COL.LAST_GENERATED).setValue(nextDue);

    // 次回生成日を計算して更新
    const newNextDue = tmpl.ecEventLabel
      ? getNextEcEventDateFromSheet(tmpl.ecEventLabel, nextDue, ecEvents)
      : computeNextRepeatDate(tmpl.startDate, tmpl.repeat, nextDue);
    if (newNextDue) autoSheet.getRange(tmpl._row, AUTO_COL.NEXT_DUE).setValue(newNextDue);
  });
}

function getNextEcEventDateFromSheet(ecEventLabel, afterDate, ecEvents) {
  if (!ecEvents || !ecEvents.length) return null;
  const after = new Date(afterDate); after.setHours(0,0,0,0);
  const baseLabel = ecEventLabel.replace('×5,0日','');
  const futures = ecEvents
    .filter(function(e) { return e.start && (e.label === ecEventLabel || e.label === baseLabel); })
    .map(function(e) { var d = new Date(e.start); d.setHours(0,0,0,0); return d; })
    .filter(function(d) { return d > after; })
    .sort(function(a,b) { return a - b; });
  return futures[0] || null;
}

function computeNextRepeatDate(startDate, repeat, afterDate) {
  const ref = new Date(afterDate); ref.setHours(0,0,0,0);
  let d = new Date(startDate); d.setHours(0,0,0,0);
  while (d <= ref) {
    if (repeat === 'daily')        d.setDate(d.getDate() + 1);
    else if (repeat === 'weekly')  d.setDate(d.getDate() + 7);
    else if (repeat === 'biweekly') d.setDate(d.getDate() + 14);
    else if (repeat === 'monthly') d.setMonth(d.getMonth() + 1);
    else break;
  }
  return d > ref ? d : null;
}

function subtractBusinessDaysGas(date, days) {
  const d = new Date(date);
  let count = 0;
  while (count < days) {
    d.setDate(d.getDate() - 1);
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) count++;
  }
  return d;
}

// トリガー設定（GASエディタから1度だけ手動実行）
function setupAutoTemplatesTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'runAutoTemplates'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runAutoTemplates')
    .timeBased().atHour(9).everyDays(1).inTimezone('Asia/Tokyo').create();
  Logger.log('オートメーション自動実行トリガーを設定しました（毎朝9時）');
}

// 時間トリガーのセットアップ（Apps Script エディタから一度だけ手動実行）
function setupMinutesTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scanMeetingMinutes')
    .forEach(t => ScriptApp.deleteTrigger(t));
  // 1時間ごとに実行
  ScriptApp.newTrigger('scanMeetingMinutes')
    .timeBased()
    .everyHours(1)
    .create();
}
