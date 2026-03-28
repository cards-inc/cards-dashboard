/************************************************************
 * チャット稼働状況をシートに書き出す（ダッシュボード連携用）
 *
 * ▼概要
 * - runInactivityAlertV2() の実行後に呼ぶ、または単独トリガーで実行
 * - 「Slack連携」シートを読み取り、各チャットの最終更新日時を取得
 * - 結果を「チャット監視」シートに書き出す
 * - ダッシュボードはGViz APIでこのシートを読む
 *
 * ▼出力シート「チャット監視」のカラム
 *   A:ストア名 / B:ツール / C:最終更新日時 / D:未稼働営業日数
 *   E:ステータス / F:担当者 / G:更新日時
 ************************************************************/

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

    // 担当者名（F列 = index 5）
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

  // 「チャット監視」シートに書き出し
  const OUTPUT_SHEET_NAME = "チャット監視";
  let outSh = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!outSh) {
    outSh = ss.insertSheet(OUTPUT_SHEET_NAME);
  }

  // ヘッダー + データ
  const header = ["ストア名", "ツール", "最終更新日時", "未稼働営業日数", "ステータス", "担当者", "取得日時"];
  outSh.clearContents();
  outSh.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length > 0) {
    outSh.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  // 書式設定
  outSh.getRange(1, 1, 1, header.length).setFontWeight("bold");
  if (rows.length > 0) {
    outSh.getRange(2, 3, rows.length, 1).setNumberFormat("yyyy/MM/dd HH:mm");
    outSh.getRange(2, 7, rows.length, 1).setNumberFormat("yyyy/MM/dd HH:mm");
  }

  Logger.log(`チャット監視シート更新完了: ${rows.length}件`);
}
