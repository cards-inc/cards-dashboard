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

// タスクシートのGID（URLの #gid= の値）
const TASK_SHEET_GID = 235656541;

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

// GIDでシートを取得する
function getSheetByGid(gid) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheet  = sheets.find(s => s.getSheetId() === Number(gid));
  if (!sheet) throw new Error('GID ' + gid + ' のシートが見つかりません');
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

  postToSlack(text, null);
}

// トリガー設定（GASエディタから1度だけ手動実行）
function setupDailyAlertTrigger() {
  // 既存トリガー削除
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyAlert')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 毎朝9時（JST）に実行
  ScriptApp.newTrigger('sendDailyAlert')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();

  Logger.log('毎朝9時のアラートトリガーを設定しました');
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
