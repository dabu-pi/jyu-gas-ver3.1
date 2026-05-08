/****************************************************
 * JREC-01 Cloud Run 疎通確認用スモークテスト関数
 *
 * 使い方:
 *   Apps Script エディタで関数を選択して「実行」
 *   結果は実行ログ（Ctrl+Enter）または Logger.log で確認
 *
 * この2関数は GAS メニューには登録しない（開発者専用）
 ****************************************************/

/**
 * スモークテスト① — /health 疎通確認
 *
 * 期待結果: HTTP 200 / {"status":"ok"}
 * 失敗時:   エラーアラート + Logger.log に詳細
 */
function V3TR_smokeHealth() {
  var props    = PropertiesService.getScriptProperties();
  var endpoint = (props.getProperty("APPGEN_ENDPOINT") || "").trim();

  if (!endpoint) {
    SpreadsheetApp.getUi().alert(
      "[smokeHealth] APPGEN_ENDPOINT が未設定です。\n" +
      "Apps Script エディタ > プロジェクトの設定 > スクリプトプロパティ を確認してください。"
    );
    return;
  }

  Logger.log("[smokeHealth] endpoint: " + endpoint);

  var resp, code, body;
  try {
    resp = UrlFetchApp.fetch(endpoint + "/health", {
      method: "get",
      muteHttpExceptions: true
    });
    code = resp.getResponseCode();
    body = resp.getContentText();
  } catch (e) {
    var msg = "[smokeHealth] 接続失敗: " + e.message;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  Logger.log("[smokeHealth] HTTP " + code + " / body: " + body);

  if (code === 200) {
    SpreadsheetApp.getUi().alert(
      "✅ /health OK\n\nHTTP: " + code + "\nBody: " + body
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "❌ /health 失敗\n\nHTTP: " + code + "\nBody: " + body.slice(0, 300)
    );
  }
}


/**
 * スモークテスト② — /generate 疎通確認（患者0件の最小 NDJSON）
 *
 * 目的: 認証・スキーマ処理・レスポンス形式を確認する
 *       実際の患者データは送らない（来院データ不要）
 *
 * 期待結果: HTTP 200 / {"status":"ok","patients":[],...}
 * 失敗例:
 *   HTTP 401 → APPGEN_SECRET と Cloud Run SECRET_KEY が不一致
 *   HTTP 400 → NDJSON フォーマット不正
 *   HTTP 500 → Cloud Run 側の処理エラー（Cloud Logging を確認）
 */
function V3TR_smokeGenerate() {
  var props     = PropertiesService.getScriptProperties();
  var endpoint  = (props.getProperty("APPGEN_ENDPOINT") || "").trim();
  var secretKey = (props.getProperty("APPGEN_SECRET")   || "").trim();

  if (!endpoint) {
    SpreadsheetApp.getUi().alert(
      "[smokeGenerate] APPGEN_ENDPOINT が未設定です。"
    );
    return;
  }
  if (!secretKey) {
    SpreadsheetApp.getUi().alert(
      "[smokeGenerate] APPGEN_SECRET が未設定です。"
    );
    return;
  }

  // 患者0件の最小 NDJSON（meta行のみ）
  var testMonth = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
  var minimalNdjson = JSON.stringify({
    _meta: true,
    schemaVersion: "3.0",
    generatedAt: Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ssXXX"),
    month: testMonth,
    patientCount: 0
  });

  var payload = JSON.stringify({ ndjson: minimalNdjson, month: testMonth });

  Logger.log("[smokeGenerate] endpoint: " + endpoint + "/generate");
  Logger.log("[smokeGenerate] payload:  " + payload);

  var resp, code, body;
  try {
    resp = UrlFetchApp.fetch(endpoint + "/generate", {
      method: "post",
      contentType: "application/json",
      headers: { "X-Secret-Key": secretKey },
      payload: payload,
      muteHttpExceptions: true
    });
    code = resp.getResponseCode();
    body = resp.getContentText();
  } catch (e) {
    var msg = "[smokeGenerate] 接続失敗: " + e.message;
    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  Logger.log("[smokeGenerate] HTTP " + code + " / body: " + body);

  if (code === 200) {
    SpreadsheetApp.getUi().alert(
      "✅ /generate OK\n\nHTTP: " + code + "\nBody: " + body.slice(0, 400)
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "❌ /generate 失敗\n\nHTTP: " + code + "\nBody: " + body.slice(0, 400) +
      "\n\n確認先:\n" +
      "  401 → APPGEN_SECRET と Cloud Run SECRET_KEY が一致しているか確認\n" +
      "  400 → NDJSON フォーマット確認\n" +
      "  500 → GCP Console > Cloud Logging でスタックトレースを確認"
    );
  }
}
