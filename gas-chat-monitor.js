/************************************************************
 * チャット稼働状況をシートに書き出す（ダッシュボード連携用）
 *
 * ▼使い方
 * - runInactivityAlertV2 がある既存のGASプロジェクトにこの関数だけ追加
 * - 共通関数（mustGetProp_, normalizeTool_, slackGetLastActivity... 等）は
 *   既存コードにあるのでそのまま使える
 * - トリガー設定: runAlertAndWriteMonitor を定期実行に設定
 *
 * ▼出力先
 * - ダッシュボードのスプレッドシート（DASHBOARD_SHEET_ID）に
 *   「チャット監視」シートを作成・更新
 *
 * ▼出力カラム
 *   A:ストア名 / B:ツール / C:最終更新日時 / D:未稼働営業日数
 *   E:ステータス / F:担当者 / G:取得日時
 ************************************************************/

// ★ダッシュボードのスプレッドシートID
const DASHBOARD_SHEET_ID = "1I8p4IVrMTEKHRdfLwwpW6jVty5N2DIlT7riHC0TW-xw";

function writeChatMonitorSheet() {
  const props = PropertiesService.getScriptProperties();

  const sheetName = mustGetProp_(props, "MASTER_SHEET_NAME");
  const slackToken = mustGetProp_(props, "SLACK_BOT_TOKEN");
  const chatworkToken = props.getProperty("CHATWORK_TOKEN") || "";

  // 読み取り元: 既存の「Slack連携」シート（このGASに紐づくスプレッドシート）
  const srcSs = SpreadsheetApp.getActiveSpreadsheet();
  const sh = srcSs.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet not found: ${sheetName}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 3) return;

  const DATA_START_ROW = 3;
  const COL_STORE_NAME = 2;  // B
  const COL_CONSULT = 6;     // F（担当者名）
  const COL_TOOL = 12;       // L
  const COL_CHAT_ID = 13;    // M

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

  // 書き出し先: ダッシュボードのスプレッドシート
  const dstSs = SpreadsheetApp.openById(DASHBOARD_SHEET_ID);
  const OUTPUT_SHEET_NAME = "チャット監視";
  let outSh = dstSs.getSheetByName(OUTPUT_SHEET_NAME);
  if (!outSh) {
    outSh = dstSs.insertSheet(OUTPUT_SHEET_NAME);
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

/** アラート通知 + シート書き出しを一括実行（トリガーにはこれを設定） */
function runAlertAndWriteMonitor() {
  runInactivityAlertV2();
  writeChatMonitorSheet();
}
