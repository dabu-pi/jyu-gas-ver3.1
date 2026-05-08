/****************************************************
 * 患者選択UI（PatientPicker）
 *
 * 患者マスタの情報を使った検索可能なプルダウンを
 * UIシートB2に設定する。
 *
 * 表示形式: "P0001｜平山 克｜ヒラヤマ カツ｜1950-01-01"
 * 内部ID:   C2 = SPLIT(B2,"｜") の1列目で自動抽出
 *
 * ★既存コードとの棲み分け
 * - UI.patientId を "C2" に変更済み（Ver3_core.js）
 * - 既存ロジックは C2 の患者IDを参照する
 * - B2 は表示用文字列（ユーザー操作用）
 * - B3 は VLOOKUP で氏名自動表示
 ****************************************************/

/** ===== 定数 ===== */
var PP_ = {
  masterSheet: "患者マスタ",
  uiSheet:     "患者画面",
  displayCol:  "検索用",       // 患者マスタ末尾に追加する列名
  separator:   "｜",           // 全角パイプ（検索時に区切りが見やすい）
};

/**
 * 患者選択UIの一括セットアップ（メニューから実行）
 *
 * 1. 患者マスタに「検索用」列を追加（末尾）
 * 2. 検索用列に表示用文字列を生成
 * 3. UI!B2 にデータバリデーション（プルダウン）を設定
 * 4. UI!C2 に患者ID抽出式を設定
 * 5. UI!B3 に氏名VLOOKUP式を設定
 */
function setupPatientPicker_V3() {
  var ss = SpreadsheetApp.getActive();
  var masterSh = ss.getSheetByName(PP_.masterSheet);
  var uiSh     = ss.getSheetByName(PP_.uiSheet);
  if (!masterSh) throw new Error(PP_.masterSheet + " シートが見つかりません");
  if (!uiSh)     throw new Error(PP_.uiSheet + " シートが見つかりません");

  // ===== Step 1: 検索用列の追加（既存なら再利用） =====
  var displayColIdx = PatientPicker_ensureDisplayCol_(masterSh);

  // ===== Step 2: 表示用文字列を生成 =====
  PatientPicker_refreshDisplayCol_(masterSh, displayColIdx);

  // ===== Step 3: B2 にプルダウン設定 =====
  PatientPicker_applyValidation_(masterSh, uiSh, displayColIdx);

  // ===== Step 4: C2 に患者ID抽出式 =====
  var sep = PP_.separator;
  // =IFERROR(TRIM(LEFT(B2, FIND("｜",B2)-1)), "")
  uiSh.getRange("C2").setFormula(
    '=IFERROR(TRIM(LEFT(B2,FIND("' + sep + '",B2)-1)),"")'
  );

  // ===== Step 5: B3 に氏名VLOOKUP式 =====
  // C2の患者IDで患者マスタA列を検索し、B列（氏名）を返す
  uiSh.getRange("B3").setFormula(
    '=IFERROR(VLOOKUP(C2,' + PP_.masterSheet + '!A:B,2,FALSE),"")'
  );

  SpreadsheetApp.getUi().alert(
    "患者選択UIセットアップ完了\n\n" +
    "・患者マスタ「" + PP_.displayCol + "」列に検索用文字列を生成\n" +
    "・B2: 患者検索プルダウン\n" +
    "・C2: 患者ID自動抽出（数式）\n" +
    "・B3: 氏名自動表示（VLOOKUP）"
  );
}

/**
 * 患者マスタの検索用列を更新（患者追加時に実行）
 */
function refreshPatientPicker_V3() {
  var ss = SpreadsheetApp.getActive();
  var masterSh = ss.getSheetByName(PP_.masterSheet);
  if (!masterSh) throw new Error(PP_.masterSheet + " シートが見つかりません");

  var displayColIdx = PatientPicker_findDisplayCol_(masterSh);
  if (displayColIdx < 1) {
    SpreadsheetApp.getUi().alert("「" + PP_.displayCol + "」列が見つかりません。先にセットアップを実行してください。");
    return;
  }

  PatientPicker_refreshDisplayCol_(masterSh, displayColIdx);

  // バリデーションも更新（行数が変わった場合に対応）
  var uiSh = ss.getSheetByName(PP_.uiSheet);
  if (uiSh) PatientPicker_applyValidation_(masterSh, uiSh, displayColIdx);

  SpreadsheetApp.getUi().alert("患者検索プルダウンを更新しました。");
}

// ============================================================
// 内部関数（PatientPicker_ プレフィックス）
// ============================================================

/** 検索用列のインデックスを探す（1-based, 見つからなければ0） */
function PatientPicker_findDisplayCol_(masterSh) {
  var lastCol = masterSh.getLastColumn();
  if (lastCol < 1) return 0;
  var headers = masterSh.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === PP_.displayCol) return i + 1;
  }
  return 0;
}

/** 検索用列がなければ末尾に追加。列インデックス(1-based)を返す */
function PatientPicker_ensureDisplayCol_(masterSh) {
  var existing = PatientPicker_findDisplayCol_(masterSh);
  if (existing > 0) return existing;

  var newCol = masterSh.getLastColumn() + 1;
  masterSh.getRange(1, newCol).setValue(PP_.displayCol);
  return newCol;
}

/** 検索用列に表示用文字列を書き込む */
function PatientPicker_refreshDisplayCol_(masterSh, displayColIdx) {
  var lastRow = masterSh.getLastRow();
  if (lastRow < 2) return;

  var headerMap = {};
  var headers = masterSh.getRange(1, 1, 1, masterSh.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || "").trim();
    if (h) headerMap[h] = i; // 0-based
  }

  var pidIdx  = headerMap["患者ID"];
  var nameIdx = headerMap["氏名"];
  var furiIdx = headerMap["フリガナ"];
  var dobIdx  = headerMap["生年月日"];

  if (pidIdx == null || nameIdx == null) {
    throw new Error("患者マスタに「患者ID」「氏名」列が必要です");
  }

  var dataRows = masterSh.getRange(2, 1, lastRow - 1, masterSh.getLastColumn()).getValues();
  var output = [];
  var sep = PP_.separator;

  for (var r = 0; r < dataRows.length; r++) {
    var pid  = String(dataRows[r][pidIdx] || "").trim();
    if (!pid) { output.push([""]); continue; }

    var name = String(dataRows[r][nameIdx] || "").trim();
    var furi = (furiIdx != null) ? String(dataRows[r][furiIdx] || "").trim() : "";
    var dob  = "";
    if (dobIdx != null && dataRows[r][dobIdx]) {
      var d = dataRows[r][dobIdx];
      if (d instanceof Date) {
        dob = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
      } else {
        dob = String(d);
      }
    }

    // 表示形式: "P0001｜平山 克｜ヒラヤマ カツ｜1950-01-01"
    var parts = [pid, name];
    if (furi) parts.push(furi);
    if (dob)  parts.push(dob);
    output.push([parts.join(sep)]);
  }

  masterSh.getRange(2, displayColIdx, output.length, 1).setValues(output);
}

/** B2にデータバリデーション（プルダウン）を設定 */
function PatientPicker_applyValidation_(masterSh, uiSh, displayColIdx) {
  var lastRow = masterSh.getLastRow();
  if (lastRow < 2) return;

  var range = masterSh.getRange(2, displayColIdx, lastRow - 1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(true)  // 自由入力も許可（部分一致検索のため）
    .build();

  uiSh.getRange("B2").setDataValidation(rule);
}
