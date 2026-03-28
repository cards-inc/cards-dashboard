/************************************************************
 * チャット未稼働アラート＋ダッシュボード連携【完全版・単体動作】
 *
 * ▼仕様
 * - 監視対象：Slack / Chatwork / Google Chat
 * - 「Slack連携」シートを参照
 *   B:ストア名 / G,I,K:担当SlackID / L:ツール / M:チャットID
 * - 最終更新から「土日祝を除いた3営業日以上」で #check_alert に通知
 * - 担当者（G/I/K列）にメンション
 * - ★上原を毎回必ずメンション
 * - 取得エラーも同じ通知にまとめる（落ちない）
 * - 同一対象は1日1回まで（Script Propertiesに記録）
 * - Slackで「投稿へのスレッド返信」も更新とみなす
 * - 稼働状況を「チャット監視」シートに書き出し（ダッシュボード連携）
 *
 * ▼Script Properties（必須）
 * - MASTER_SHEET_NAME   : Slack連携
 * - SLACK_BOT_TOKEN     : xoxb-...
 * - SLACK_ALERT_CHANNEL : check_alert（チャンネル名）
 *
 * ▼Script Properties（任意・推奨）
 * - SLACK_ALERT_CHANNEL_ID : Cxxxx
 * - CHATWORK_TOKEN         : Chatwork監視時のみ必須
 ************************************************************/

/* ========================================================
 * メイン：Slackアラート通知
 * ======================================================== */
function runInactivityAlertV2() {
  const props = PropertiesService.getScriptProperties();

  const sheetName = mustGetProp_(props, "MASTER_SHEET_NAME");
  const slackToken = mustGetProp_(props, "SLACK_BOT_TOKEN");
  const slackAlertChannelName = mustGetProp_(props, "SLACK_ALERT_CHANNEL");
  const slackAlertChannelIdProp = props.getProperty("SLACK_ALERT_CHANNEL_ID") || "";
  const chatworkToken = props.getProperty("CHATWORK_TOKEN") || "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet not found: ${sheetName}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 3) return;

  const DATA_START_ROW = 3;
  const COL_STORE_NAME = 2;
  const COL_CONSULT_SID = 7;
  const COL_AD_SID = 9;
  const COL_ASST_SID = 11;
  const COL_TOOL = 12;
  const COL_CHAT_ID = 13;

  const ALWAYS_MENTION_IDS = ["U08BB4MH8P3"];

  const now = new Date();
  const todayKey = formatYmd_(now);

  const alertChannelId = slackAlertChannelIdProp
    ? slackAlertChannelIdProp
    : slackResolveChannelIdByName_(slackToken, slackAlertChannelName);

  if (!alertChannelId) throw new Error(`Slack alert channel not found: #${slackAlertChannelName}`);

  const alerts = [];

  for (let r = DATA_START_ROW; r <= values.length; r++) {
    const row = values[r - 1];

    const storeName = String(row[COL_STORE_NAME - 1] || "").trim();
    const toolRaw = String(row[COL_TOOL - 1] || "").trim();
    const chatIdRaw = String(row[COL_CHAT_ID - 1] || "").trim();
    if (!storeName || !toolRaw || !chatIdRaw) continue;

    const slackIds = uniq_([
      ...ALWAYS_MENTION_IDS.map(cleanSlackId_),
      cleanSlackId_(row[COL_CONSULT_SID - 1]),
      cleanSlackId_(row[COL_AD_SID - 1]),
      cleanSlackId_(row[COL_ASST_SID - 1]),
    ].filter(Boolean));

    const mention = slackIds.length ? slackIds.map(id => `<@${id}>`).join(" ") : "";

    let lastActivity = null;
    let errorNote = "";

    try {
      const tool = normalizeTool_(toolRaw);

      if (tool === "Slack") {
        lastActivity = slackGetLastActivityTimeIncludingThreads_(slackToken, chatIdRaw);
      } else if (tool === "Chatwork") {
        if (!chatworkToken) throw new Error("CHATWORK_TOKEN is missing");
        lastActivity = chatworkGetLastUpdateTime_(chatworkToken, chatIdRaw);
      } else if (tool === "GoogleChat") {
        lastActivity = gchatGetLastMessageTime_(chatIdRaw);
      } else {
        throw new Error(`Unknown tool: ${toolRaw}`);
      }
    } catch (e) {
      errorNote = String(e && e.message ? e.message : e);
    }

    if (!lastActivity) {
      if (errorNote) {
        alerts.push({
          status: "ERROR",
          store: storeName,
          tool: toolRaw,
          chatId: chatIdRaw,
          mentions: mention,
          detail: errorNote,
        });
      }
      continue;
    }

    const bizDays = businessDaysBetween_(lastActivity, now);

    if (bizDays >= 3) {
      const key = `notified_${storeName}_${toolRaw}_${chatIdRaw}`;
      if (props.getProperty(key) === todayKey) continue;
      props.setProperty(key, todayKey);

      alerts.push({
        status: "INACTIVE",
        store: storeName,
        tool: toolRaw,
        chatId: chatIdRaw,
        mentions: mention,
        last: lastActivity,
        biz: bizDays,
      });
    }
  }

  if (alerts.length === 0) return;

  const inactive = alerts.filter(a => a.status === "INACTIVE").sort((a, b) => b.biz - a.biz);
  const errors = alerts.filter(a => a.status === "ERROR");

  const lines = [];

  if (inactive.length) {
    lines.push("🚨 *3営業日以上動いていないチャットがあります*");
    lines.push("");

    inactive.forEach(a => {
      lines.push(
        `【${a.store}】\n` +
        `ツール：${a.tool}\n` +
        `担当：${a.mentions || "—"}\n` +
        `最終更新：${formatYmdHm_(a.last)}（${a.biz}営業日未更新）`
      );
      lines.push("");
    });
  }

  if (errors.length) {
    lines.push("⚠️ *取得エラー（設定/権限を要確認）*");
    lines.push("");

    errors.forEach(e => {
      lines.push(
        `【${e.store}】\n` +
        `ツール：${e.tool}\n` +
        `担当：${e.mentions || "—"}\n` +
        `エラー内容：${e.detail}`
      );
      lines.push("");
    });
  }

  if (lines.length) {
    slackChatPostMessage_(slackToken, alertChannelId, lines.join("\n"));
  }
}

/* ========================================================
 * メイン：ダッシュボード用シート書き出し
 * ======================================================== */
function writeChatMonitorSheet() {
  const props = PropertiesService.getScriptProperties();

  const sheetName = mustGetProp_(props, "MASTER_SHEET_NAME");
  const slackToken = mustGetProp_(props, "SLACK_BOT_TOKEN");
  const chatworkToken = props.getProperty("CHATWORK_TOKEN") || "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet not found: ${sheetName}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 3) return;

  const DATA_START_ROW = 3;
  const COL_STORE_NAME = 2;
  const COL_CONSULT = 6;
  const COL_TOOL = 12;
  const COL_CHAT_ID = 13;

  const now = new Date();
  const rows = [];

  for (let r = DATA_START_ROW; r <= values.length; r++) {
    const row = values[r - 1];

    const storeName = String(row[COL_STORE_NAME - 1] || "").trim();
    const toolRaw = String(row[COL_TOOL - 1] || "").trim();
    const chatIdRaw = String(row[COL_CHAT_ID - 1] || "").trim();
    if (!storeName || !toolRaw || !chatIdRaw) continue;

    const assignee = String(row[COL_CONSULT - 1] || "").trim();

    let lastActivity = null;
    let errorNote = "";

    try {
      const tool = normalizeTool_(toolRaw);

      if (tool === "Slack") {
        lastActivity = slackGetLastActivityTimeIncludingThreads_(slackToken, chatIdRaw);
      } else if (tool === "Chatwork") {
        if (!chatworkToken) throw new Error("CHATWORK_TOKEN is missing");
        lastActivity = chatworkGetLastUpdateTime_(chatworkToken, chatIdRaw);
      } else if (tool === "GoogleChat") {
        lastActivity = gchatGetLastMessageTime_(chatIdRaw);
      } else {
        throw new Error(`Unknown tool: ${toolRaw}`);
      }
    } catch (e) {
      errorNote = String(e && e.message ? e.message : e);
    }

    let bizDays = 0;
    let status = "";

    if (lastActivity) {
      bizDays = businessDaysBetween_(lastActivity, now);
      if (bizDays >= 5) {
        status = "危険";
      } else if (bizDays >= 3) {
        status = "警告";
      } else {
        status = "正常";
      }
    } else {
      status = errorNote ? "エラー" : "データなし";
    }

    rows.push([
      storeName,
      toolRaw,
      lastActivity || "",
      bizDays,
      status,
      assignee,
      now,
    ]);
  }

  const OUTPUT_SHEET_NAME = "チャット監視";
  let outSh = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!outSh) {
    outSh = ss.insertSheet(OUTPUT_SHEET_NAME);
  }

  const header = ["ストア名", "ツール", "最終更新日時", "未稼働営業日数", "ステータス", "担当者", "取得日時"];
  outSh.clearContents();
  outSh.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length > 0) {
    outSh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  outSh.getRange(1, 1, 1, header.length).setFontWeight("bold");
  if (rows.length > 0) {
    outSh.getRange(2, 3, rows.length, 1).setNumberFormat("yyyy/MM/dd HH:mm");
    outSh.getRange(2, 7, rows.length, 1).setNumberFormat("yyyy/MM/dd HH:mm");
  }

  Logger.log(`チャット監視シート更新完了: ${rows.length}件`);
}

/** アラート通知 + シート書き出しを一括実行 */
function runAlertAndWriteMonitor() {
  runInactivityAlertV2();
  writeChatMonitorSheet();
}

/* ================= Slack ================= */

function slackResolveChannelIdByName_(token, channelName) {
  const channels = slackListAllChannels_(token);
  const found = channels.find(c => c.name === channelName);
  return found ? found.id : null;
}

function slackListAllChannels_(token) {
  const results = [];
  let cursor = "";
  while (true) {
    const url = "https://slack.com/api/conversations.list"
      + "?types=public_channel,private_channel"
      + "&exclude_archived=true"
      + "&limit=200"
      + (cursor ? `&cursor=${encodeURIComponent(cursor)}` : "");

    const body = slackApiGet_(token, url);
    (body.channels || []).forEach(c => results.push({ id: c.id, name: c.name }));
    cursor = body.response_metadata?.next_cursor || "";
    if (!cursor) break;
  }
  return results;
}

function slackGetLastActivityTimeIncludingThreads_(token, channelId) {
  const historyUrl = "https://slack.com/api/conversations.history"
    + `?channel=${encodeURIComponent(channelId)}`
    + "&limit=50";

  const history = slackApiGet_(token, historyUrl);
  const msgs = history.messages || [];
  if (!msgs.length) return null;

  let maxTsNum = 0;

  for (const m of msgs) {
    if (m.ts) {
      const tsNum = Number(m.ts);
      if (!isNaN(tsNum) && tsNum > maxTsNum) maxTsNum = tsNum;
    }

    if (m.thread_ts && (m.reply_count || 0) > 0) {
      const threadMaxTsNum = slackGetThreadLatestTsNum_(token, channelId, m.thread_ts);
      if (threadMaxTsNum && threadMaxTsNum > maxTsNum) maxTsNum = threadMaxTsNum;
    }
  }

  if (!maxTsNum) return null;
  const sec = Math.floor(maxTsNum);
  return new Date(sec * 1000);
}

function slackGetThreadLatestTsNum_(token, channelId, threadTs) {
  const url = "https://slack.com/api/conversations.replies"
    + `?channel=${encodeURIComponent(channelId)}`
    + `&ts=${encodeURIComponent(threadTs)}`
    + "&limit=100";

  const body = slackApiGet_(token, url);
  const msgs = body.messages || [];
  if (!msgs.length) return 0;

  let maxTsNum = 0;
  for (const m of msgs) {
    if (!m.ts) continue;
    const tsNum = Number(m.ts);
    if (!isNaN(tsNum) && tsNum > maxTsNum) maxTsNum = tsNum;
  }
  return maxTsNum;
}

function slackApiGet_(token, url) {
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  const text = res.getContentText() || "{}";
  const json = JSON.parse(text);
  if (!json.ok) throw new Error(`Slack API error: ${res.getResponseCode()} ${text}`);
  return json;
}

function slackChatPostMessage_(token, channelId, text) {
  const payload = { channel: channelId, text };
  const res = UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method: "post",
    contentType: "application/json; charset=utf-8",
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || "{}");
  if (!body.ok) throw new Error(`chat.postMessage failed: ${res.getResponseCode()} ${res.getContentText()}`);
}

/* ================= Chatwork ================= */

function chatworkGetLastUpdateTime_(token, roomIdRaw) {
  const roomId = String(roomIdRaw).trim();
  const url = `https://api.chatwork.com/v2/rooms/${encodeURIComponent(roomId)}`;
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "x-chatworktoken": token },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) throw new Error(`Chatwork API error: ${code} ${body}`);
  const json = JSON.parse(body);
  if (!json.last_update_time) return null;
  return new Date(Number(json.last_update_time) * 1000);
}

/* ================= Google Chat ================= */

function gchatGetLastMessageTime_(chatIdRaw) {
  const space = normalizeGchatSpace_(chatIdRaw);
  const url = `https://chat.googleapis.com/v1/${space}/messages?pageSize=1&orderBy=createTime%20desc`;
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) throw new Error(`Google Chat API error: ${code} ${body}`);
  const json = JSON.parse(body);
  const msgs = json.messages || [];
  if (!msgs.length) return null;
  const ct = msgs[0].createTime || msgs[0].updateTime;
  if (!ct) return null;
  return new Date(ct);
}

function normalizeGchatSpace_(raw) {
  const v = String(raw).trim();
  if (v.startsWith("spaces/")) return v;
  return `spaces/${v}`;
}

/* ================= 営業日計算（祝日内蔵） ================= */

function businessDaysBetween_(start, end) {
  const s = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const e = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  if (e <= s) return 0;

  let days = 0;
  const d = new Date(s);

  while (d < e) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (dow === 0 || dow === 6) continue;
    if (isJapanHoliday_(d)) continue;
    days++;
  }
  return days;
}

function isJapanHoliday_(dateObj) {
  const y = dateObj.getFullYear();
  const key = formatYmd_(dateObj);

  const base = japanHolidayBaseSet_(y);
  if (base.has(key)) return true;

  const substitute = japanSubstituteHolidays_(y, base);
  if (substitute.has(key)) return true;

  const citizens = japanCitizensHolidays_(y, base, substitute);
  if (citizens.has(key)) return true;

  return false;
}

function japanHolidayBaseSet_(year) {
  const set = new Set();

  add_(set, year, 1, 1);
  add_(set, year, 2, 11);
  add_(set, year, 2, 23);
  add_(set, year, 4, 29);
  add_(set, year, 5, 3);
  add_(set, year, 5, 4);
  add_(set, year, 5, 5);
  add_(set, year, 8, 11);
  add_(set, year, 11, 3);
  add_(set, year, 11, 23);

  addNthMonday_(set, year, 1, 2);
  addNthMonday_(set, year, 7, 3);
  addNthMonday_(set, year, 9, 3);
  addNthMonday_(set, year, 10, 2);

  addSpringEquinox_(set, year);
  addAutumnEquinox_(set, year);

  return set;
}

function japanSubstituteHolidays_(year, baseSet) {
  const out = new Set();
  const start = new Date(year, 0, 1);
  const end = new Date(year + 1, 0, 1);
  const d = new Date(start);

  while (d < end) {
    const key = formatYmd_(d);
    if (baseSet.has(key) && d.getDay() === 0) {
      const x = new Date(d);
      while (true) {
        x.setDate(x.getDate() + 1);
        const xdow = x.getDay();
        if (xdow === 0 || xdow === 6) continue;
        const xkey = formatYmd_(x);
        if (baseSet.has(xkey)) continue;
        if (out.has(xkey)) continue;
        out.add(xkey);
        break;
      }
    }
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function japanCitizensHolidays_(year, baseSet, substituteSet) {
  const out = new Set();
  const start = new Date(year, 0, 1);
  const end = new Date(year + 1, 0, 1);
  const d = new Date(start);

  while (d < end) {
    if (![0, 6].includes(d.getDay())) {
      const key = formatYmd_(d);
      const already = baseSet.has(key) || substituteSet.has(key);
      if (!already) {
        const prev = new Date(d); prev.setDate(prev.getDate() - 1);
        const next = new Date(d); next.setDate(next.getDate() + 1);
        const prevKey = formatYmd_(prev);
        const nextKey = formatYmd_(next);

        const prevHoliday = baseSet.has(prevKey) || substituteSet.has(prevKey);
        const nextHoliday = baseSet.has(nextKey) || substituteSet.has(nextKey);
        if (prevHoliday && nextHoliday) out.add(key);
      }
    }
    d.setDate(d.getDate() + 1);
  }
  return out;
}

/* ================= ユーティリティ ================= */

function normalizeTool_(toolRaw) {
  const t = String(toolRaw || "").trim().toLowerCase().replace(/\s+/g, "");
  if (t === "slack") return "Slack";
  if (t === "chatwork") return "Chatwork";
  if (t.includes("googlechat") || t.includes("gchat")) return "GoogleChat";
  if (t.includes("google") && t.includes("chat")) return "GoogleChat";
  return toolRaw;
}

function uniq_(arr) {
  return Array.from(new Set(arr));
}

function cleanSlackId_(v) {
  const s = String(v || "").trim();
  return /^[UW][A-Z0-9]+$/i.test(s) ? s : "";
}

function formatYmd_(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${dd}`;
}

function formatYmdHm_(d) {
  const ymd = formatYmd_(d);
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  return `${ymd} ${hh}:${mm}`;
}

function mustGetProp_(props, key) {
  const v = props.getProperty(key);
  if (!v) throw new Error(`Missing script property: ${key}`);
  return v;
}

function add_(set, y, m, d) {
  set.add(`${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`);
}
function addNthMonday_(set, y, m, nth) {
  const first = new Date(y, m - 1, 1);
  const firstDow = first.getDay();
  const offsetToMon = (8 - firstDow) % 7;
  const day = 1 + offsetToMon + (nth - 1) * 7;
  add_(set, y, m, day);
}
function addSpringEquinox_(set, y) {
  const day = Math.floor(20.8431 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  add_(set, y, 3, day);
}
function addAutumnEquinox_(set, y) {
  const day = Math.floor(23.2488 + 0.242194 * (y - 1980) - Math.floor((y - 1980) / 4));
  add_(set, y, 9, day);
}
