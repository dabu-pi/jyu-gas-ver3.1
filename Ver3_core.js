/****************************************************
 * 柔整 Ver3.1 統合版（保存＋金額＋ヘッダ追記を一括実行）
 *
 * ★主要変更（統合版）
 * - saveVisit_V3() で ①来院ケース保存 → ②金額計算 → ③来院ヘッダ追記 を一括
 * - visitKey = 患者ID + "_" + yyyy-MM-dd（旧"|"から変更）
 * - 同日二重登録禁止（来院ヘッダに同一visitKeyがあればエラー停止）
 * - 上書き禁止（修正は別機能で行う前提）
 *
 * ★既存仕様（維持）
 * - 30日ルール（エピソード連結＋終了境界で打ち切り）
 * - 安全弁（コアなし行のチェック強制FALSE等）
 * - 終了日セルに何か入っていれば終了扱い
 * - 終了は「コア補完を止める」だけ（治療チェックは消さない）
 *
 * ★金額計算
 * - 来院ケースの部位データから直接算定
 * - 来院合計 = 初検料 + 再検料 + 相談支援料 + 明細合計
 * - 窓口負担額 = 来院合計 × 負担割合（roundToUnit_で丸め）
 * - 保険請求額 = 来院合計 - 窓口負担額（丸めない）
 ****************************************************/

/** ===== シート名 ===== */
const SHEETS = {
  settings: "設定",
  cases: "来院ケース",
  master: "患者マスタ",
  ui: "患者画面",
  detail: "施術明細",
  header: "来院ヘッダ",
  history: "初検情報履歴",
  insurer: "保険者情報",
};

/** ===== 患者画面 UIセル ===== */
const UI = {
  patientId: "B2",
  treatDate: "B4",
  kubun: "B5",

  // 表示専用：区分（ユーザー指定）
  case1_kubunView: "C10",
  case2_kubunView: "C25",

  // Case1（入力2行）※ A〜G（G=終了日）
  case1_rows: ["A12:G12", "A13:G13"],
  // G列を終了日に使うため、所見/経過/履歴は1列右へ（H〜）
  case1_shoken: "H11:M13",
  case1_keikaNow: "H16:M17",
  case1_keikaHistory: "H19:M23",

  // Case2（入力2行）※ A〜G（G=終了日）
  case2_rows: ["A27:G27", "A28:G28"],
  case2_shoken: "H26:M28",
  case2_keikaNow: "H31:M32",
  case2_keikaHistory: "H34:M38",

  // 終了日入力（明示）
  case1_endHeader: "G11",
  case2_endHeader: "G26",
};

/** ===== 来院ケース列名（誤解ゼロ命名：部位1/2） ===== */
const CASE_COLS = {
  visitKey: "visitKey",
  treatDate: "施術日",
  patientId: "患者ID",
  caseNo: "caseNo",
  injuryFixed: "受傷日_確定",
  kubun: "区分",
  initFee: "初検料",
  reFee: "再検料",
  supportFee: "相談支援料",
  detailSum: "明細合計(case)",
  caseTotal: "case合計",
  createdAt: "作成日時",
  caseKey: "caseKey",

  // 部位1
  p1: "部位_部位1",
  d1: "傷病_部位1",
  inj1: "受傷日_部位1",
  cold1: "冷罨法_部位1",
  warm1: "温罨法_部位1",
  elec1: "電療_部位1",

  // 部位2
  p2: "部位_部位2",
  d2: "傷病_部位2",
  inj2: "受傷日_部位2",
  cold2: "冷罨法_部位2",
  warm2: "温罨法_部位2",
  elec2: "電療_部位2",

  // 施術開始日/終了日（部位ごと）
  start1: "施術開始日_部位1",
  end1: "施術終了日_部位1",
  start2: "施術開始日_部位2",
  end2: "施術終了日_部位2",

  shoken: "所見",
  keikaNow: "経過_今回",
};

/** ===== 来院ヘッダ列名 ===== */
const HEADER_COLS = {
  visitKey: "visitKey",
  treatDate: "施術日",
  patientId: "患者ID",
  kubun: "区分",
  injuryVisit: "受傷日_確定(来院)",
  initFee: "初検料",
  reFee: "再検料",
  supportFee: "相談支援料",
  detailSum: "明細合計",
  visitTotal: "来院合計",
  lastVisit: "最終来院日",
  gapDays: "前回から日数",
  needCheck: "要確認",
  createdAt: "作成日時",
  windowPay: "窓口負担額",
  claimPay: "保険請求額",
  caseKey: "caseKey",
  caseIndex: "caseIndex",
};

/** ===== 患者マスタ列名 ===== */
const MASTER_COLS = {
  patientId: "患者ID",
  burden: "負担割合",
};



/** ===== メニュー ===== */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("柔整ツール")
      .addItem("来院ケースへ保存（統合：保存＋金額＋ヘッダ）", "saveVisit_V3")
      .addItem("経過履歴を更新（患者画面）", "refreshKeikaHistoryUI_V3")
      .addItem("自動引継ぎを実行（2回目以降）", "autofillFromPreviousVisit_V3")
      .addSeparator()
      .addItem("来院ケース → 来院ヘッダへ出力（高速）", "exportHeaderFromCases_V3")
      .addSeparator()
      .addItem("患者画面クリア（入力のみ）", "clearEntryUI_V3")
      .addSeparator()
      .addItem("ヘッダー確認（デバッグ）", "checkHeaders_V3")
      .addItem("金額再計算（施術明細→ヘッダ）", "menuRecalcAmounts_V3")
      .addItem("申請書_転記データ作成（患者×月）", "V3TR_menuBuildTransferData")
      .addToUi();
  } catch (err) {
    console.error(err);
  }
}

/** ===== onEdit ===== */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    var sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.ui) return;

    var a1 = e.range.getA1Notation();
    if (a1 !== UI.patientId && a1 !== UI.treatDate) return;

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(500)) return;

    try {
      var props = PropertiesService.getScriptProperties();
      var now = Date.now();
      var last = Number(props.getProperty("V3_ONEDIT_LAST") || 0);
      if (now - last < 1200) return;
      props.setProperty("V3_ONEDIT_LAST", String(now));

      refreshKeikaHistoryUI_V3();
      autofillFromPreviousVisit_V3();
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
  }
}

/** ===== ヘッダー名正規化（trim＋全角空白→半角＋連続空白圧縮＋英字小文字） ===== */
function normalizeHeaderName_(s) {
  return String(s || "")
    .trim()
    .replace(/\u3000/g, " ")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

/** ===== 別名辞書（正式名 → 別名リスト） ===== */
const HEADER_ALIASES_ = {
  "患者ID":    ["患者ＩＤ", "patientId", "PATIENT_ID", "patient_id"],
  "施術日":    ["施術日付", "treatDate", "treatmentDate", "treatment_date"],
  "visitKey":  ["訪問キー", "来院キー", "キー", "visitkey"],
  "区分":      ["kubun", "種別"],
  "受傷日_確定": ["受傷日", "負傷日", "injuryDate", "injury_date"],
  "受傷日_部位1": ["受傷日1"],
  "受傷日_部位2": ["受傷日2"],
  "傷病_部位1": ["傷病1", "傷病名_部位1"],
  "傷病_部位2": ["傷病2", "傷病名_部位2"],
  "部位_部位1": ["部位1", "患部_部位1", "site1"],
  "部位_部位2": ["部位2", "患部_部位2", "site2"],
  "caseNo":    ["caseno", "ケース番号"],
  "caseKey":   ["casekey", "ケースキー"],
};

var _aliasToCanonical_ = null;
function getAliasToCanonicalMap_() {
  if (_aliasToCanonical_) return _aliasToCanonical_;
  _aliasToCanonical_ = {};
  for (var canonical in HEADER_ALIASES_) {
    var aliases = HEADER_ALIASES_[canonical];
    for (var j = 0; j < aliases.length; j++) {
      _aliasToCanonical_[normalizeHeaderName_(aliases[j])] = canonical;
    }
  }
  return _aliasToCanonical_;
}

/** ===== 1行目ヘッダー名→列番号（1-based）＋正規化＋別名吸収 ===== */
function buildHeaderColMap_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var aliasMap = getAliasToCanonicalMap_();
  var map = {};

  headers.forEach(function(raw, i) {
    var trimmed = String(raw || "").trim();
    if (!trimmed) return;

    var col1 = i + 1;

    if (!map[trimmed]) map[trimmed] = col1;

    var norm = normalizeHeaderName_(trimmed);
    if (norm !== trimmed && !map[norm]) map[norm] = col1;

    var canonical = aliasMap[norm];
    if (canonical && !map[canonical]) map[canonical] = col1;
  });

  return map;
}

/** ===== 結合セル（左上） ===== */
function getMergedValue_(sheet, a1Range) {
  return sheet.getRange(a1Range).getCell(1, 1).getValue();
}
function setMergedValue_(sheet, a1Range, value) {
  var cell = sheet.getRange(a1Range).getCell(1, 1);
  cell.setValue(value);
  cell.setWrap(true);
}

/** ===== 日付ユーティリティ ===== */
function fmt_(d, pat) {
  return Utilities.formatDate(d, "Asia/Tokyo", pat);
}
function daysBetween_(fromDate, toDate) {
  var a = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  var b = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
  return Math.round((b.getTime() - a.getTime()) / (24 * 3600 * 1000));
}
function minDate_(d1, d2) {
  var a = (d1 instanceof Date) ? d1 : null;
  var b = (d2 instanceof Date) ? d2 : null;
  if (!a && !b) return "";
  if (a && !b) return a;
  if (!a && b) return b;
  return (a.getTime() <= b.getTime()) ? a : b;
}

/** ===== visitKey / caseKey 構築 ===== */
function buildVisitKey_(patientId, treatDate) {
  return patientId + "_" + fmt_(treatDate, "yyyy-MM-dd");
}
function buildCaseKey_(visitKey, caseNo) {
  return visitKey + "_C" + caseNo;
}

/** ===== 必須列チェック（シート名・実ヘッダー付きエラー） ===== */
function ensureRequiredCols_(map, requiredList, sheetName) {
  var missing = requiredList.filter(function(n) { return !map[n]; });
  if (!missing.length) return;

  var actualHeaders = Object.keys(map).slice(0, 30);
  var label = sheetName ? ("【" + sheetName + "】") : "【対象シート】";

  throw new Error(
    label + " ヘッダー不足：\n" +
    "- " + missing.join("\n- ") + "\n\n" +
    label + " 実ヘッダー（先頭30件）：\n" +
    actualHeaders.join(", ")
  );
}

/** ===== キー列で行検索 ===== */
function findRowByKey_(sheet, map, keyHeaderName, keyValue) {
  var c = map[keyHeaderName];
  if (!c) throw new Error("キー列がありません: " + keyHeaderName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var vals = sheet.getRange(2, c, lastRow - 1, 1).getValues().flat();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i] || "").trim() === String(keyValue)) return i + 2;
  }
  return 0;
}

/** ===== rowArrへ列名でセット ===== */
function setByName_(rowArr, headerMap, name, value, opt) {
  opt = opt || {};
  var col = headerMap[name];
  if (!col) throw new Error("対象シートに列がありません: " + name);
  var idx = col - 1;
  if (opt.preserveIfExists && rowArr[idx] !== "" && rowArr[idx] != null) return;
  rowArr[idx] = value;
}

/** ===== チェック安全弁 ===== */
function coreHasAny_(rowVals) {
  var part = String(rowVals[0] || "").trim();
  var dis  = String(rowVals[1] || "").trim();
  var inj  = rowVals[2] instanceof Date;
  return !!(part || dis || inj);
}
function forceChecksFalse_(rowVals) {
  rowVals[3] = false;
  rowVals[4] = false;
  rowVals[5] = false;
}
function forceWarmElecTrue_(rowVals) {
  rowVals[4] = true;
  rowVals[5] = true;
}

/** ===== 終了判定 ===== */
function isEnded_(endVal, treatDate) {
  if (endVal === "" || endVal == null) return false;
  if (!(endVal instanceof Date)) return true;
  if (!(treatDate instanceof Date)) return true;
  return endVal.getTime() <= treatDate.getTime();
}

/** ===== UIの2行入力を読む ===== */
function readRowNewUI_(uiSh, a1) {
  var v = uiSh.getRange(a1).getValues()[0];

  var part = String(v[0] || "").trim();
  var disease = String(v[1] || "").trim();
  var injuryDate = (v[2] instanceof Date) ? v[2] : "";

  var cold = v[3] === true;
  var warm = v[4] === true;
  var elec = v[5] === true;

  var endRaw = v[6];
  var endVal =
    (endRaw instanceof Date) ? endRaw :
    (String(endRaw || "").trim() ? String(endRaw) : "");

  var hasCore = !!(part || disease || injuryDate);

  return { part: part, disease: disease, injuryDate: injuryDate, cold: cold, warm: warm, elec: elec, endVal: endVal, hasCore: hasCore };
}

/** ===== 来院ケースから患者の来院日一覧 ===== */
function getPatientVisitDatesFromCases_(caseSh, caseMap, patientId) {
  var lastRow = caseSh.getLastRow();
  if (lastRow < 2) return [];

  var n = lastRow - 1;
  var cPid = caseMap[CASE_COLS.patientId];
  var cDt  = caseMap[CASE_COLS.treatDate];

  var pidVals = caseSh.getRange(2, cPid, n, 1).getValues().flat();
  var dtVals  = caseSh.getRange(2, cDt,  n, 1).getValues().flat();

  var set = new Map();
  for (var i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    var d = dtVals[i];
    if (!(d instanceof Date)) continue;
    var k = fmt_(d, "yyyy-MM-dd");
    if (!set.has(k)) set.set(k, d);
  }
  return Array.from(set.values()).sort(function(a,b){ return a.getTime()-b.getTime(); });
}


/* =======================================================================
   saveVisit_V3  ―  統合保存（ケース保存＋金額計算＋ヘッダ追記）
   ======================================================================= */
function saveVisit_V3() {
  var ss = SpreadsheetApp.getActive();
  var uiSh = ss.getSheetByName(SHEETS.ui);
  var caseSh = ss.getSheetByName(SHEETS.cases);
  var headSh = ss.getSheetByName(SHEETS.header);
  if (!uiSh || !caseSh || !headSh) throw new Error("必要シートが見つかりません（患者画面/来院ケース/来院ヘッダ）");

  var patientId = String(uiSh.getRange(UI.patientId).getValue() || "").trim();
  var treatDate = uiSh.getRange(UI.treatDate).getValue();

  if (!patientId) throw new Error("患者ID（B2）が空です。");
  if (!(treatDate instanceof Date)) throw new Error("来院日（B4）が日付になっていません。");

  var visitKey = buildVisitKey_(patientId, treatDate);
  var now = new Date();

  var caseMap = buildHeaderColMap_(caseSh);
  var headMap = buildHeaderColMap_(headSh);
  ensureRequiredCols_(caseMap, Object.values(CASE_COLS), SHEETS.cases);
  ensureRequiredCols_(headMap, Object.values(HEADER_COLS), SHEETS.header);

  // ★二重登録禁止チェック（来院ヘッダに同一visitKeyがあれば停止）
  var existingHeaderRow = findRowByKey_(headSh, headMap, HEADER_COLS.visitKey, visitKey);
  if (existingHeaderRow > 0) {
    throw new Error(
      "同日二重登録禁止：来院ヘッダに同一visitKeyが存在します。\n" +
      "visitKey: " + visitKey + "（行 " + existingHeaderRow + "）\n" +
      "修正は別機能で行ってください。"
    );
  }

  // ① 来院ケースへ保存
  var ep1 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 1);
  var ep2 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 2);

  upsertOneCase_(uiSh, caseSh, caseMap, {
    visitKey: visitKey, patientId: patientId, treatDate: treatDate,
    kubun: ep1.kubun,
    caseNo: 1,
    now: now,
    episodeStartDate: ep1.episodeStartDate
  });

  upsertOneCase_(uiSh, caseSh, caseMap, {
    visitKey: visitKey, patientId: patientId, treatDate: treatDate,
    kubun: ep2.kubun,
    caseNo: 2,
    now: now,
    episodeStartDate: ep2.episodeStartDate
  });

  // ② 金額計算（来院ケースベース）
  var amounts = calcHeaderAmountsByVisitKey_V3_(ss, visitKey, patientId, treatDate, ep1.kubun, ep2.kubun);

  // ③ 来院ヘッダへ1行追記
  var injuryFixed = null;
  var caseKey1 = buildCaseKey_(visitKey, 1);
  var row1idx = findRowByKey_(caseSh, caseMap, CASE_COLS.caseKey, caseKey1);
  if (row1idx > 0) {
    var v = caseSh.getRange(row1idx, caseMap[CASE_COLS.injuryFixed], 1, 1).getValue();
    if (v instanceof Date) injuryFixed = v;
  }
  if (!injuryFixed) {
    var caseKey2 = buildCaseKey_(visitKey, 2);
    var row2idx = findRowByKey_(caseSh, caseMap, CASE_COLS.caseKey, caseKey2);
    if (row2idx > 0) {
      var v2 = caseSh.getRange(row2idx, caseMap[CASE_COLS.injuryFixed], 1, 1).getValue();
      if (v2 instanceof Date) injuryFixed = v2;
    }
  }

  var lastVisit = findLastVisitDateInHeader_(headSh, headMap, patientId, treatDate);
  var gapDays = (lastVisit instanceof Date) ? daysBetween_(lastVisit, treatDate) : "";

  var kubunLabel = ep1.kubun || ep2.kubun || "";

  appendHeaderRow_V3_(headSh, headMap, {
    visitKey: visitKey,
    treatDate: treatDate,
    patientId: patientId,
    kubun: kubunLabel,
    injuryVisit: injuryFixed,
    initFee: amounts.initFee,
    reFee: amounts.reFee,
    supportFee: amounts.supportFee,
    detailSum: amounts.detailSum,
    visitTotal: amounts.visitTotal,
    windowPay: amounts.windowPay,
    claimPay: amounts.claimPay,
    lastVisit: lastVisit || "",
    gapDays: gapDays,
    needCheck: (gapDays !== "" && gapDays > 30) ? "要確認" : "",
    createdAt: now,
    caseKey: buildCaseKey_(visitKey, 1),
    caseIndex: 1
  });

  refreshKeikaHistoryUI_V3();
  clearAfterSaveUI_V3_(uiSh);
  SpreadsheetApp.getUi().alert(
    "保存完了（統合：ケース＋金額＋ヘッダ）\n" +
    "visitKey: " + visitKey + "\n" +
    "来院合計: " + amounts.visitTotal + "\n" +
    "窓口負担: " + amounts.windowPay + "\n" +
    "保険請求: " + amounts.claimPay
  );
}


/* =======================================================================
   calcHeaderAmountsByVisitKey_V3_  ―  金額計算（来院ケースベース）
   ======================================================================= */
function calcHeaderAmountsByVisitKey_V3_(ss, visitKey, patientId, treatDate, kubun1, kubun2) {
  var settings = loadSettings_V3_(ss);

  var masterSh = ss.getSheetByName(SHEETS.master);
  var masterMap = buildHeaderColMap_(masterSh);
  var burden = loadBurdenRatio_V3_(masterSh, masterMap, patientId);

  var caseSh = ss.getSheetByName(SHEETS.cases);
  var caseMap = buildHeaderColMap_(caseSh);

  var hasInit = (kubun1 === "初検" || kubun2 === "初検");
  var hasRe   = (kubun1 === "再検" || kubun2 === "再検");

  var initFee    = hasInit ? settings.initFee : 0;
  var reFee      = hasRe ? settings.reFee : 0;
  var supportFee = hasInit ? settings.initSupport : 0;

  var detail1 = calcCaseDetailAmount_(caseSh, caseMap, visitKey, 1, kubun1, treatDate, settings);
  var detail2 = calcCaseDetailAmount_(caseSh, caseMap, visitKey, 2, kubun2, treatDate, settings);
  var detailSum = detail1 + detail2;

  var visitTotal = initFee + reFee + supportFee + detailSum;

  var unit = settings.roundUnit || 1;
  var windowPay = roundToUnit_V3_(visitTotal * burden, unit);

  var claimPay = visitTotal - windowPay;

  return {
    initFee: initFee,
    reFee: reFee,
    supportFee: supportFee,
    detailSum: detailSum,
    visitTotal: visitTotal,
    windowPay: windowPay,
    claimPay: claimPay
  };
}

/** 1ケース分の明細金額を来院ケースの部位データから算定 */
function calcCaseDetailAmount_(caseSh, caseMap, visitKey, caseNo, kubun, treatDate, settings) {
  var caseKey = buildCaseKey_(visitKey, caseNo);
  var rowIndex = findRowByKey_(caseSh, caseMap, CASE_COLS.caseKey, caseKey);
  if (rowIndex === 0) return 0;

  var row = caseSh.getRange(rowIndex, 1, 1, caseSh.getLastColumn()).getValues()[0];
  var get = function(name) { return row[caseMap[name] - 1]; };

  var total = 0;
  var partCount = 0;

  // 部位1
  var p1 = String(get(CASE_COLS.p1) || "").trim();
  var d1 = String(get(CASE_COLS.d1) || "").trim();
  var inj1 = get(CASE_COLS.inj1);
  if (p1 || d1 || (inj1 instanceof Date)) {
    partCount++;
    total += calcOnePartAmount_(settings, kubun, d1, inj1, treatDate,
      get(CASE_COLS.cold1) === true,
      get(CASE_COLS.warm1) === true,
      get(CASE_COLS.elec1) === true,
      partCount);
  }

  // 部位2
  var p2 = String(get(CASE_COLS.p2) || "").trim();
  var d2 = String(get(CASE_COLS.d2) || "").trim();
  var inj2 = get(CASE_COLS.inj2);
  if (p2 || d2 || (inj2 instanceof Date)) {
    partCount++;
    total += calcOnePartAmount_(settings, kubun, d2, inj2, treatDate,
      get(CASE_COLS.cold2) === true,
      get(CASE_COLS.warm2) === true,
      get(CASE_COLS.elec2) === true,
      partCount);
  }

  return total;
}

/** 1部位分の金額算定（厚労省運用ルール準拠） */
function calcOnePartAmount_(settings, kubun, byomei, injuryDate, treatDate, coldChk, warmChk, elecChk, partOrder) {
  var injuryType = detectInjuryType_V3_(byomei);
  var base = calcBaseFee_V3_(settings, kubun, injuryType);

  var dayDiff = null;
  if (injuryDate instanceof Date && treatDate instanceof Date) {
    dayDiff = daysBetween_(injuryDate, treatDate);
  }

  var cold = (coldChk && kubun === "初検" && dayDiff != null && dayDiff <= 1) ? settings.cold : 0;
  var warm = (warmChk && (kubun === "再検" || kubun === "後療") && dayDiff != null && dayDiff >= 5) ? settings.warm : 0;
  var electro = (elecChk && (kubun === "再検" || kubun === "後療") && dayDiff != null && dayDiff >= 5) ? settings.electro : 0;
  var taiki = ((warm > 0 || electro > 0) && (kubun === "再検" || kubun === "後療") && dayDiff != null && dayDiff >= 5) ? settings.taiki : 0;

  var coef = (partOrder >= 3) ? Number(settings.multiCoef3 || 0.6) : 1.0;

  return (base + cold + warm + electro + taiki) * coef;
}


/* =======================================================================
   appendHeaderRow_V3_  ―  来院ヘッダへ1行追記（visitKey重複時throw）
   ======================================================================= */
function appendHeaderRow_V3_(headSh, headMap, obj) {
  // 二重登録禁止チェック（念のため再チェック）
  var existingRow = findRowByKey_(headSh, headMap, HEADER_COLS.visitKey, obj.visitKey);
  if (existingRow > 0) {
    throw new Error(
      "来院ヘッダに同一visitKeyが存在します（二重登録禁止）。\n" +
      "visitKey: " + obj.visitKey + "（行 " + existingRow + "）"
    );
  }

  var rowArr = new Array(headSh.getLastColumn()).fill("");

  setByName_(rowArr, headMap, HEADER_COLS.visitKey, obj.visitKey);
  setByName_(rowArr, headMap, HEADER_COLS.treatDate, obj.treatDate);
  setByName_(rowArr, headMap, HEADER_COLS.patientId, obj.patientId);
  setByName_(rowArr, headMap, HEADER_COLS.kubun, obj.kubun);
  if (obj.injuryVisit instanceof Date) {
    setByName_(rowArr, headMap, HEADER_COLS.injuryVisit, obj.injuryVisit);
  }
  setByName_(rowArr, headMap, HEADER_COLS.initFee, obj.initFee);
  setByName_(rowArr, headMap, HEADER_COLS.reFee, obj.reFee);
  setByName_(rowArr, headMap, HEADER_COLS.supportFee, obj.supportFee);
  setByName_(rowArr, headMap, HEADER_COLS.detailSum, obj.detailSum);
  setByName_(rowArr, headMap, HEADER_COLS.visitTotal, obj.visitTotal);
  setByName_(rowArr, headMap, HEADER_COLS.windowPay, obj.windowPay);
  setByName_(rowArr, headMap, HEADER_COLS.claimPay, obj.claimPay);
  setByName_(rowArr, headMap, HEADER_COLS.lastVisit, obj.lastVisit);
  setByName_(rowArr, headMap, HEADER_COLS.gapDays, obj.gapDays);
  setByName_(rowArr, headMap, HEADER_COLS.needCheck, obj.needCheck);
  setByName_(rowArr, headMap, HEADER_COLS.createdAt, obj.createdAt);
  setByName_(rowArr, headMap, HEADER_COLS.caseKey, obj.caseKey);
  setByName_(rowArr, headMap, HEADER_COLS.caseIndex, obj.caseIndex);

  headSh.getRange(headSh.getLastRow() + 1, 1, 1, headSh.getLastColumn()).setValues([rowArr]);
}


/* =====================================================
   upsertOneCase_（来院ケースへ1ケース保存）
   ===================================================== */
function upsertOneCase_(uiSh, caseSh, caseMap, base) {
  var visitKey = base.visitKey;
  var patientId = base.patientId;
  var treatDate = base.treatDate;
  var kubun = base.kubun;
  var caseNo = base.caseNo;
  var now = base.now;
  var episodeStartDate = base.episodeStartDate;

  var rows = (caseNo === 1) ? UI.case1_rows : UI.case2_rows;
  var shokenRange = (caseNo === 1) ? UI.case1_shoken : UI.case2_shoken;
  var keikaRange  = (caseNo === 1) ? UI.case1_keikaNow : UI.case2_keikaNow;

  var line1 = readRowNewUI_(uiSh, rows[0]);
  var line2 = readRowNewUI_(uiSh, rows[1]);

  var shoken = String(getMergedValue_(uiSh, shokenRange) || "").trim();
  var keikaNow = String(getMergedValue_(uiSh, keikaRange) || "").trim();

  var hasAny = line1.hasCore || line2.hasCore || !!shoken || !!keikaNow || !!line1.endVal || !!line2.endVal;
  if (!hasAny) return;

  var injuryFixed = minDate_(line1.injuryDate, line2.injuryDate);
  var caseKey = buildCaseKey_(visitKey, caseNo);
  var rowIndex = findRowByKey_(caseSh, caseMap, CASE_COLS.caseKey, caseKey);

  if (rowIndex === 0) {
    var rowArr = new Array(caseSh.getLastColumn()).fill("");

    setByName_(rowArr, caseMap, CASE_COLS.visitKey, visitKey);
    setByName_(rowArr, caseMap, CASE_COLS.treatDate, treatDate);
    setByName_(rowArr, caseMap, CASE_COLS.patientId, patientId);
    setByName_(rowArr, caseMap, CASE_COLS.caseNo, caseNo);
    setByName_(rowArr, caseMap, CASE_COLS.caseKey, caseKey);
    setByName_(rowArr, caseMap, CASE_COLS.kubun, kubun);

    if (injuryFixed) setByName_(rowArr, caseMap, CASE_COLS.injuryFixed, injuryFixed);

    writeLinesToCaseRow_(rowArr, caseMap, line1, line2);

    if (line1.hasCore) setByName_(rowArr, caseMap, CASE_COLS.start1, (kubun === "初検") ? treatDate : episodeStartDate);
    if (line2.hasCore) setByName_(rowArr, caseMap, CASE_COLS.start2, (kubun === "初検") ? treatDate : episodeStartDate);

    if (line1.endVal !== "" && line1.endVal != null) setByName_(rowArr, caseMap, CASE_COLS.end1, line1.endVal);
    if (line2.endVal !== "" && line2.endVal != null) setByName_(rowArr, caseMap, CASE_COLS.end2, line2.endVal);

    if (shoken) setByName_(rowArr, caseMap, CASE_COLS.shoken, shoken);
    if (keikaNow) setByName_(rowArr, caseMap, CASE_COLS.keikaNow, keikaNow);

    setByName_(rowArr, caseMap, CASE_COLS.initFee, "");
    setByName_(rowArr, caseMap, CASE_COLS.reFee, "");
    setByName_(rowArr, caseMap, CASE_COLS.supportFee, "");
    setByName_(rowArr, caseMap, CASE_COLS.detailSum, "");
    setByName_(rowArr, caseMap, CASE_COLS.caseTotal, "");
    setByName_(rowArr, caseMap, CASE_COLS.createdAt, now);

    caseSh.appendRow(rowArr);
  } else {
    var lastCol = caseSh.getLastColumn();
    var rowArr2 = caseSh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

    setByName_(rowArr2, caseMap, CASE_COLS.visitKey, visitKey);
    setByName_(rowArr2, caseMap, CASE_COLS.treatDate, treatDate);
    setByName_(rowArr2, caseMap, CASE_COLS.patientId, patientId);
    setByName_(rowArr2, caseMap, CASE_COLS.caseNo, caseNo);
    setByName_(rowArr2, caseMap, CASE_COLS.caseKey, caseKey);
    setByName_(rowArr2, caseMap, CASE_COLS.kubun, kubun);

    if (injuryFixed) setByName_(rowArr2, caseMap, CASE_COLS.injuryFixed, injuryFixed, { preserveIfExists: true });

    writeLinesToCaseRow_(rowArr2, caseMap, line1, line2);

    if (line1.hasCore) {
      if (kubun === "初検") {
        setByName_(rowArr2, caseMap, CASE_COLS.start1, treatDate);
      } else {
        setByName_(rowArr2, caseMap, CASE_COLS.start1, episodeStartDate, { preserveIfExists: true });
      }
    }
    if (line2.hasCore) {
      if (kubun === "初検") {
        setByName_(rowArr2, caseMap, CASE_COLS.start2, treatDate);
      } else {
        setByName_(rowArr2, caseMap, CASE_COLS.start2, episodeStartDate, { preserveIfExists: true });
      }
    }

    if (line1.endVal !== "" && line1.endVal != null) setByName_(rowArr2, caseMap, CASE_COLS.end1, line1.endVal);
    if (line2.endVal !== "" && line2.endVal != null) setByName_(rowArr2, caseMap, CASE_COLS.end2, line2.endVal);

    if (shoken) setByName_(rowArr2, caseMap, CASE_COLS.shoken, shoken);
    if (keikaNow) setByName_(rowArr2, caseMap, CASE_COLS.keikaNow, keikaNow);

    caseSh.getRange(rowIndex, 1, 1, lastCol).setValues([rowArr2]);
  }
}

function writeLinesToCaseRow_(rowArr, caseMap, line1, line2) {
  setByName_(rowArr, caseMap, CASE_COLS.p1, line1.part || "");
  setByName_(rowArr, caseMap, CASE_COLS.d1, line1.disease || "");
  setByName_(rowArr, caseMap, CASE_COLS.inj1, line1.injuryDate || "");
  setByName_(rowArr, caseMap, CASE_COLS.cold1, line1.cold);
  setByName_(rowArr, caseMap, CASE_COLS.warm1, line1.warm);
  setByName_(rowArr, caseMap, CASE_COLS.elec1, line1.elec);

  setByName_(rowArr, caseMap, CASE_COLS.p2, line2.part || "");
  setByName_(rowArr, caseMap, CASE_COLS.d2, line2.disease || "");
  setByName_(rowArr, caseMap, CASE_COLS.inj2, line2.injuryDate || "");
  setByName_(rowArr, caseMap, CASE_COLS.cold2, line2.cold);
  setByName_(rowArr, caseMap, CASE_COLS.warm2, line2.warm);
  setByName_(rowArr, caseMap, CASE_COLS.elec2, line2.elec);
}

/** ===== 経過履歴（最新5件） ===== */
function refreshKeikaHistoryUI_V3() {
  var ss = SpreadsheetApp.getActive();
  var uiSh = ss.getSheetByName(SHEETS.ui);
  var caseSh = ss.getSheetByName(SHEETS.cases);
  if (!uiSh || !caseSh) return;

  var patientId = String(uiSh.getRange(UI.patientId).getValue() || "").trim();
  if (!patientId) {
    setMergedValue_(uiSh, UI.case1_keikaHistory, "");
    setMergedValue_(uiSh, UI.case2_keikaHistory, "");
    return;
  }

  var caseMap = buildHeaderColMap_(caseSh);
  ensureRequiredCols_(caseMap, [CASE_COLS.patientId, CASE_COLS.caseNo, CASE_COLS.treatDate, CASE_COLS.keikaNow], SHEETS.cases);

  var hist1 = buildKeikaHistoryTextFromCases_(caseSh, caseMap, patientId, 1, 5);
  var hist2 = buildKeikaHistoryTextFromCases_(caseSh, caseMap, patientId, 2, 5);

  setMergedValue_(uiSh, UI.case1_keikaHistory, hist1);
  setMergedValue_(uiSh, UI.case2_keikaHistory, hist2);
}

function buildKeikaHistoryTextFromCases_(caseSh, caseMap, patientId, caseNo, limit) {
  var cPid = caseMap[CASE_COLS.patientId];
  var cNo  = caseMap[CASE_COLS.caseNo];
  var cDt  = caseMap[CASE_COLS.treatDate];
  var cK   = caseMap[CASE_COLS.keikaNow];

  var lastRow = caseSh.getLastRow();
  if (lastRow < 2) return "";

  var pidVals = caseSh.getRange(2, cPid, lastRow - 1, 1).getValues().flat();
  var noVals  = caseSh.getRange(2, cNo,  lastRow - 1, 1).getValues().flat();
  var dtVals  = caseSh.getRange(2, cDt,  lastRow - 1, 1).getValues().flat();
  var kVals   = caseSh.getRange(2, cK,   lastRow - 1, 1).getValues().flat();

  var rows = [];
  for (var i = 0; i < pidVals.length; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    if (Number(noVals[i] || 0) !== caseNo) continue;

    var d = dtVals[i];
    var kv = String(kVals[i] || "").trim();
    if (!(d instanceof Date)) continue;
    if (!kv) continue;
    rows.push({ d: d, k: kv });
  }

  rows.sort(function(a, b) { return b.d.getTime() - a.d.getTime(); });
  return rows.slice(0, limit).map(function(x) { return fmt_(x.d, "M/d") + "：" + x.k; }).join("\n");
}

/** ===== 自動引継ぎ ===== */
function autofillFromPreviousVisit_V3() {
  var ss = SpreadsheetApp.getActive();
  var uiSh = ss.getSheetByName(SHEETS.ui);
  var caseSh = ss.getSheetByName(SHEETS.cases);
  if (!uiSh || !caseSh) return;

  var patientId = String(uiSh.getRange(UI.patientId).getValue() || "").trim();
  var treatDate = uiSh.getRange(UI.treatDate).getValue();

  if (!patientId || !(treatDate instanceof Date)) {
    uiSh.getRange(UI.case1_kubunView).setValue("");
    uiSh.getRange(UI.case2_kubunView).setValue("");
    return;
  }

  var caseMap = buildHeaderColMap_(caseSh);
  ensureRequiredCols_(caseMap, [
    CASE_COLS.patientId, CASE_COLS.treatDate, CASE_COLS.caseNo,
    CASE_COLS.p1, CASE_COLS.d1, CASE_COLS.inj1,
    CASE_COLS.p2, CASE_COLS.d2, CASE_COLS.inj2,
    CASE_COLS.start1, CASE_COLS.end1,
    CASE_COLS.start2, CASE_COLS.end2,
    CASE_COLS.shoken
  ], SHEETS.cases);

  var ep1 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 1);
  var ep2 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 2);

  var latest1 = findLatestCaseRowDateInEpisode_(caseSh, caseMap, patientId, treatDate, 1);
  var latest2 = findLatestCaseRowDateInEpisode_(caseSh, caseMap, patientId, treatDate, 2);

  var src1 = latest1 ? findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, latest1, 1) : null;
  var src2 = latest2 ? findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, latest2, 2) : null;

  // case1
  (function() {
    var caseClosed = isCaseClosedAsOf_(src1, treatDate);
    if (caseClosed) {
      uiSh.getRange(UI.case1_kubunView).setValue("初検");
      setMergedValue_(uiSh, UI.case1_shoken, "");
      applyCaseRowToUI_Safe_(uiSh, null, 1, treatDate, { forceWarmElec: false });
    } else {
      uiSh.getRange(UI.case1_kubunView).setValue(ep1.kubun || "");
      applyCaseRowToUI_Safe_(uiSh, src1, 1, treatDate, { forceWarmElec: true });
      var curShoken = String(getMergedValue_(uiSh, UI.case1_shoken) || "").trim();
      if (!curShoken && src1) {
        var startForShoken = minDate_(src1.start1, src1.start2);
        var shokenDate = (startForShoken instanceof Date) ? startForShoken : ep1.episodeStartDate;
        var shokenSrc = findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, shokenDate, 1);
        if (shokenSrc) {
          var srcText = String(shokenSrc.shoken || "").trim();
          if (srcText) setMergedValue_(uiSh, UI.case1_shoken, srcText);
        }
      }
    }
  })();

  // case2
  (function() {
    var caseClosed = isCaseClosedAsOf_(src2, treatDate);
    if (caseClosed) {
      uiSh.getRange(UI.case2_kubunView).setValue("初検");
      setMergedValue_(uiSh, UI.case2_shoken, "");
      applyCaseRowToUI_Safe_(uiSh, null, 2, treatDate, { forceWarmElec: false });
    } else {
      uiSh.getRange(UI.case2_kubunView).setValue(ep2.kubun || "");
      applyCaseRowToUI_Safe_(uiSh, src2, 2, treatDate, { forceWarmElec: true });
      var curShoken = String(getMergedValue_(uiSh, UI.case2_shoken) || "").trim();
      if (!curShoken && src2) {
        var startForShoken = minDate_(src2.start1, src2.start2);
        var shokenDate = (startForShoken instanceof Date) ? startForShoken : ep2.episodeStartDate;
        var shokenSrc = findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, shokenDate, 2);
        if (shokenSrc) {
          var srcText = String(shokenSrc.shoken || "").trim();
          if (srcText) setMergedValue_(uiSh, UI.case2_shoken, srcText);
        }
      }
    }
  })();
}

function sameDateKey_(d) {
  return (d instanceof Date) ? fmt_(d, "yyyy-MM-dd") : "";
}

function findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, dateObj, caseNo) {
  if (!dateObj) return null;
  var lastRow = caseSh.getLastRow();
  if (lastRow < 2) return null;

  var n = lastRow - 1;
  var cPid = caseMap[CASE_COLS.patientId];
  var cDt  = caseMap[CASE_COLS.treatDate];
  var cNo  = caseMap[CASE_COLS.caseNo];

  var pidVals = caseSh.getRange(2, cPid, n, 1).getValues().flat();
  var dtVals  = caseSh.getRange(2, cDt,  n, 1).getValues().flat();
  var noVals  = caseSh.getRange(2, cNo,  n, 1).getValues().flat();

  var targetKey = sameDateKey_(dateObj);

  for (var i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    if (Number(noVals[i] || 0) !== caseNo) continue;
    var d = dtVals[i];
    if (!(d instanceof Date)) continue;
    if (sameDateKey_(d) !== targetKey) continue;

    var rowIndex = i + 2;
    var row = caseSh.getRange(rowIndex, 1, 1, caseSh.getLastColumn()).getValues()[0];
    var get = function(name) { return row[caseMap[name] - 1]; };

    return {
      p1: String(get(CASE_COLS.p1) || ""),
      d1: String(get(CASE_COLS.d1) || ""),
      inj1: (get(CASE_COLS.inj1) instanceof Date) ? get(CASE_COLS.inj1) : "",
      cold1: get(CASE_COLS.cold1) === true,
      warm1: get(CASE_COLS.warm1) === true,
      elec1: get(CASE_COLS.elec1) === true,
      start1: (get(CASE_COLS.start1) instanceof Date) ? get(CASE_COLS.start1) : "",
      end1: get(CASE_COLS.end1),

      p2: String(get(CASE_COLS.p2) || ""),
      d2: String(get(CASE_COLS.d2) || ""),
      inj2: (get(CASE_COLS.inj2) instanceof Date) ? get(CASE_COLS.inj2) : "",
      cold2: get(CASE_COLS.cold2) === true,
      warm2: get(CASE_COLS.warm2) === true,
      elec2: get(CASE_COLS.elec2) === true,
      start2: (get(CASE_COLS.start2) instanceof Date) ? get(CASE_COLS.start2) : "",
      end2: get(CASE_COLS.end2),

      shoken: String(get(CASE_COLS.shoken) || ""),
    };
  }
  return null;
}

/** ===== UIへ安全に適用 ===== */
function applyCaseRowToUI_Safe_(uiSh, src, caseNo, treatDate, opt) {
  opt = opt || {};
  var rows = (caseNo === 1) ? UI.case1_rows : UI.case2_rows;

  var rng1 = uiSh.getRange(rows[0]);
  var rng2 = uiSh.getRange(rows[1]);

  var v1 = rng1.getValues()[0];
  var v2 = rng2.getValues()[0];

  if (!src) {
    if (!coreHasAny_(v1)) forceChecksFalse_(v1);
    if (!coreHasAny_(v2)) forceChecksFalse_(v2);
    rng1.setValues([v1]);
    rng2.setValues([v2]);
    return;
  }

  var ended1 = isEnded_(src.end1, treatDate);
  var ended2 = isEnded_(src.end2, treatDate);

  var hasSrc1 = !!(String(src.p1||"").trim() || String(src.d1||"").trim() || (src.inj1 instanceof Date));
  var hasSrc2 = !!(String(src.p2||"").trim() || String(src.d2||"").trim() || (src.inj2 instanceof Date));

  if (hasSrc1 && (v1[6] === "" || v1[6] == null)) {
    if (src.end1 !== "" && src.end1 != null) v1[6] = src.end1;
  }
  if (hasSrc2 && (v2[6] === "" || v2[6] == null)) {
    if (src.end2 !== "" && src.end2 != null) v2[6] = src.end2;
  }

  if (!ended1) {
    if (!coreHasAny_(v1)) {
      v1[0] = src.p1 || "";
      v1[1] = src.d1 || "";
      v1[2] = src.inj1 || "";
    }
  }

  if (!ended2) {
    if (!coreHasAny_(v2)) {
      v2[0] = src.p2 || "";
      v2[1] = src.d2 || "";
      v2[2] = src.inj2 || "";
    }
  }

  var row1HasCore = coreHasAny_(v1);
  var row2HasCore = coreHasAny_(v2);

  if (!row1HasCore) forceChecksFalse_(v1);
  if (!row2HasCore) forceChecksFalse_(v2);

  if (opt.forceWarmElec) {
    if (row1HasCore) forceWarmElecTrue_(v1);
    if (row2HasCore) forceWarmElecTrue_(v2);
  }

  rng1.setValues([v1]);
  rng2.setValues([v2]);
}

/** ===== 来院ケース → 来院ヘッダへ一括出力（高速） ===== */
function exportHeaderFromCases_V3() {
  var ss = SpreadsheetApp.getActive();
  var caseSh = ss.getSheetByName(SHEETS.cases);
  var headSh = ss.getSheetByName(SHEETS.header);
  if (!caseSh || !headSh) throw new Error("必要シートが見つかりません（来院ケース/来院ヘッダ）");

  var caseMap = buildHeaderColMap_(caseSh);
  var headMap = buildHeaderColMap_(headSh);

  ensureRequiredCols_(caseMap, [CASE_COLS.visitKey, CASE_COLS.treatDate, CASE_COLS.patientId, CASE_COLS.kubun, CASE_COLS.injuryFixed, CASE_COLS.caseKey, CASE_COLS.caseNo], SHEETS.cases);
  ensureRequiredCols_(headMap, Object.values(HEADER_COLS), SHEETS.header);

  var lastRow = caseSh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("来院ケースにデータがありません。");
    return;
  }

  var n = lastRow - 1;
  var visitKeyVals = caseSh.getRange(2, caseMap[CASE_COLS.visitKey], n, 1).getValues().flat();
  var treatVals    = caseSh.getRange(2, caseMap[CASE_COLS.treatDate], n, 1).getValues().flat();
  var pidVals      = caseSh.getRange(2, caseMap[CASE_COLS.patientId], n, 1).getValues().flat();
  var kubunVals    = caseSh.getRange(2, caseMap[CASE_COLS.kubun], n, 1).getValues().flat();
  var injFixVals   = caseSh.getRange(2, caseMap[CASE_COLS.injuryFixed], n, 1).getValues().flat();
  var caseKeyVals  = caseSh.getRange(2, caseMap[CASE_COLS.caseKey], n, 1).getValues().flat();
  var caseNoVals   = caseSh.getRange(2, caseMap[CASE_COLS.caseNo], n, 1).getValues().flat();

  var existed = buildExistingHeaderKeySet_(headSh, headMap);

  var out = [];
  var now = new Date();

  for (var i = 0; i < n; i++) {
    var visitKey = String(visitKeyVals[i] || "").trim();
    var patientId = String(pidVals[i] || "").trim();
    var treatDate = treatVals[i];
    var kubun = String(kubunVals[i] || "").trim();
    var inj = injFixVals[i];
    var caseKey = String(caseKeyVals[i] || "").trim();
    var caseIndex = Number(caseNoVals[i] || 0);

    if (!visitKey || !patientId || !(treatDate instanceof Date) || !caseKey || !caseIndex) continue;

    var uniq = caseKey + "#" + caseIndex;
    if (existed.has(uniq)) continue;

    var rowArr = new Array(headSh.getLastColumn()).fill("");

    setByName_(rowArr, headMap, HEADER_COLS.visitKey, visitKey);
    setByName_(rowArr, headMap, HEADER_COLS.treatDate, treatDate);
    setByName_(rowArr, headMap, HEADER_COLS.patientId, patientId);
    setByName_(rowArr, headMap, HEADER_COLS.kubun, kubun);
    if (inj instanceof Date) setByName_(rowArr, headMap, HEADER_COLS.injuryVisit, inj);

    setByName_(rowArr, headMap, HEADER_COLS.initFee, "");
    setByName_(rowArr, headMap, HEADER_COLS.reFee, "");
    setByName_(rowArr, headMap, HEADER_COLS.supportFee, "");
    setByName_(rowArr, headMap, HEADER_COLS.detailSum, "");
    setByName_(rowArr, headMap, HEADER_COLS.visitTotal, "");
    setByName_(rowArr, headMap, HEADER_COLS.windowPay, "");

    var last = findLastVisitDateInHeader_(headSh, headMap, patientId, treatDate);
    setByName_(rowArr, headMap, HEADER_COLS.lastVisit, (last instanceof Date) ? last : "");
    setByName_(rowArr, headMap, HEADER_COLS.gapDays, (last instanceof Date) ? daysBetween_(last, treatDate) : "");

    setByName_(rowArr, headMap, HEADER_COLS.needCheck, "");
    setByName_(rowArr, headMap, HEADER_COLS.createdAt, now);
    setByName_(rowArr, headMap, HEADER_COLS.caseKey, caseKey);
    setByName_(rowArr, headMap, HEADER_COLS.caseIndex, caseIndex);

    out.push(rowArr);
    existed.add(uniq);
  }

  if (!out.length) {
    SpreadsheetApp.getUi().alert("出力対象がありません（すでに出力済み or データ不足）");
    return;
  }

  headSh.getRange(headSh.getLastRow() + 1, 1, out.length, headSh.getLastColumn()).setValues(out);
  SpreadsheetApp.getUi().alert("来院ヘッダへ出力しました：" + out.length + " 行");
}

function buildExistingHeaderKeySet_(headSh, headMap) {
  var set = new Set();
  var lastRow = headSh.getLastRow();
  if (lastRow < 2) return set;

  var cCaseKey = headMap[HEADER_COLS.caseKey];
  var cCaseIdx = headMap[HEADER_COLS.caseIndex];
  if (!cCaseKey || !cCaseIdx) return set;

  var keys = headSh.getRange(2, cCaseKey, lastRow - 1, 1).getValues().flat();
  var idxs = headSh.getRange(2, cCaseIdx, lastRow - 1, 1).getValues().flat();

  for (var i = 0; i < keys.length; i++) {
    var k = String(keys[i] || "").trim();
    var n = Number(idxs[i] || 0);
    if (k && n) set.add(k + "#" + n);
  }
  return set;
}

function findLastVisitDateInHeader_(headSh, headMap, patientId, treatDate) {
  var lastRow = headSh.getLastRow();
  if (lastRow < 2) return null;

  var cPid = headMap[HEADER_COLS.patientId];
  var cDt  = headMap[HEADER_COLS.treatDate];
  if (!cPid || !cDt) return null;

  var pidVals = headSh.getRange(2, cPid, lastRow - 1, 1).getValues().flat();
  var dtVals  = headSh.getRange(2, cDt,  lastRow - 1, 1).getValues().flat();

  var best = null;
  for (var i = 0; i < pidVals.length; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    var d = dtVals[i];
    if (!(d instanceof Date)) continue;
    if (d.getTime() >= treatDate.getTime()) continue;
    if (!best || d.getTime() > best.getTime()) best = d;
  }
  return best;
}

/** ===== クリア（入力だけ） ===== */
function clearEntryUI_V3() {
  var ss = SpreadsheetApp.getActive();
  var uiSh = ss.getSheetByName(SHEETS.ui);

  uiSh.getRange("B2").clearContent();
  uiSh.getRange("F2:F4").clearContent();
  uiSh.getRange("B5:B7").clearContent();

  uiSh.getRange("A12:G13").clearContent();
  uiSh.getRange("H11:M14").clearContent();
  uiSh.getRange("H16:M17").clearContent();
  uiSh.getRange("H19:M23").clearContent();

  uiSh.getRange("A27:G28").clearContent();
  uiSh.getRange("H26:M29").clearContent();
  uiSh.getRange("H31:M32").clearContent();
  uiSh.getRange("H34:M38").clearContent();

  uiSh.getRange("D12:F13").setValues([
    [false, false, false],
    [false, false, false]
  ]);
  uiSh.getRange("D27:F28").setValues([
    [false, false, false],
    [false, false, false]
  ]);

  uiSh.getRange(UI.case1_kubunView).clearContent();
  uiSh.getRange(UI.case2_kubunView).clearContent();

  SpreadsheetApp.getUi().alert("自動入力エリアをクリアしました（B4は保持・書式は保持）。");
}

/** ===== 保存後クリア ===== */
function clearAfterSaveUI_V3_(uiSh) {
  uiSh.getRange(UI.patientId).setValue("");
  uiSh.getRange(UI.kubun).setValue("");

  UI.case1_rows.forEach(function(r) { uiSh.getRange(r).clearContent(); });
  UI.case2_rows.forEach(function(r) { uiSh.getRange(r).clearContent(); });

  uiSh.getRange("D12:F13").setValues([[false,false,false],[false,false,false]]);
  uiSh.getRange("D27:F28").setValues([[false,false,false],[false,false,false]]);

  setMergedValue_(uiSh, UI.case1_shoken, "");
  setMergedValue_(uiSh, UI.case1_keikaNow, "");
  setMergedValue_(uiSh, UI.case2_shoken, "");
  setMergedValue_(uiSh, UI.case2_keikaNow, "");

  setMergedValue_(uiSh, UI.case1_keikaHistory, "");
  setMergedValue_(uiSh, UI.case2_keikaHistory, "");

  uiSh.getRange(UI.case1_kubunView).setValue("");
  uiSh.getRange(UI.case2_kubunView).setValue("");
}

/** ===== ヘッダー確認 ===== */
function checkHeaders_V3() {
  var ss = SpreadsheetApp.getActive();
  var caseSh = ss.getSheetByName(SHEETS.cases);
  var headSh = ss.getSheetByName(SHEETS.header);
  if (!caseSh || !headSh) throw new Error("来院ケース または 来院ヘッダ が見つかりません。");

  var caseMap = buildHeaderColMap_(caseSh);
  var headMap = buildHeaderColMap_(headSh);

  var needCase = Object.values(CASE_COLS);
  var needHead = Object.values(HEADER_COLS);

  var missCase = needCase.filter(function(h) { return !caseMap[h]; });
  var missHead = needHead.filter(function(h) { return !headMap[h]; });

  var caseActual = Object.keys(caseMap).slice(0, 30).join(", ");
  var headActual = Object.keys(headMap).slice(0, 30).join(", ");

  SpreadsheetApp.getUi().alert(
    "ヘッダーチェック結果\n\n" +
    (missCase.length
      ? "【来院ケース 不足】\n- " + missCase.join("\n- ") + "\n\n実ヘッダー：" + caseActual + "\n\n"
      : "【来院ケース 不足】なし\n\n") +
    (missHead.length
      ? "【来院ヘッダ 不足】\n- " + missHead.join("\n- ") + "\n\n実ヘッダー：" + headActual
      : "【来院ヘッダ 不足】なし")
  );
}

/** ===== caseNo別の来院日（コアがある日だけ） ===== */
function getPatientVisitDatesFromCasesByCase_(caseSh, caseMap, patientId, caseNo) {
  var lastRow = caseSh.getLastRow();
  if (lastRow < 2) return [];

  var n = lastRow - 1;
  var cPid = caseMap[CASE_COLS.patientId];
  var cDt  = caseMap[CASE_COLS.treatDate];
  var cNo  = caseMap[CASE_COLS.caseNo];

  var cP1 = caseMap[CASE_COLS.p1];
  var cD1 = caseMap[CASE_COLS.d1];
  var cI1 = caseMap[CASE_COLS.inj1];
  var cP2 = caseMap[CASE_COLS.p2];
  var cD2 = caseMap[CASE_COLS.d2];
  var cI2 = caseMap[CASE_COLS.inj2];

  var pidVals = caseSh.getRange(2, cPid, n, 1).getValues().flat();
  var dtVals  = caseSh.getRange(2, cDt,  n, 1).getValues().flat();
  var noVals  = caseSh.getRange(2, cNo,  n, 1).getValues().flat();

  var p1Vals = caseSh.getRange(2, cP1, n, 1).getValues().flat();
  var d1Vals = caseSh.getRange(2, cD1, n, 1).getValues().flat();
  var i1Vals = caseSh.getRange(2, cI1, n, 1).getValues().flat();
  var p2Vals = caseSh.getRange(2, cP2, n, 1).getValues().flat();
  var d2Vals = caseSh.getRange(2, cD2, n, 1).getValues().flat();
  var i2Vals = caseSh.getRange(2, cI2, n, 1).getValues().flat();

  var uniq = new Map();

  for (var i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    if (Number(noVals[i] || 0) !== caseNo) continue;

    var d = dtVals[i];
    if (!(d instanceof Date)) continue;

    var hasCore =
      String(p1Vals[i] || "").trim() ||
      String(d1Vals[i] || "").trim() ||
      (i1Vals[i] instanceof Date) ||
      String(p2Vals[i] || "").trim() ||
      String(d2Vals[i] || "").trim() ||
      (i2Vals[i] instanceof Date);

    if (!hasCore) continue;

    var key = fmt_(d, "yyyy-MM-dd");
    if (!uniq.has(key)) uniq.set(key, d);
  }

  return Array.from(uniq.values()).sort(function(a, b) { return a.getTime() - b.getTime(); });
}

/** ===== エピソード計算（30日ルール＋終了境界で打ち切り） ===== */
function calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, caseNo) {
  var dates = getPatientVisitDatesFromCasesByCase_(caseSh, caseMap, patientId, caseNo);

  if (!dates.length) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  var prevDates = dates
    .filter(function(d) { return d.getTime() < treatDate.getTime(); })
    .sort(function(a, b) { return a.getTime() - b.getTime(); });

  if (!prevDates.length) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  var lastDate = prevDates[prevDates.length - 1];
  var gap = daysBetween_(lastDate, treatDate);
  if (gap > 30) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  var episode = buildEpisodeDatesBackwards_StopAtClosed_(
    caseSh, caseMap, patientId, caseNo, prevDates, treatDate, 30
  );

  if (!episode.length) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  var episodeStartDate = episode[0];
  var priorCountInEpisode = episode.length;

  var kubun =
    (priorCountInEpisode === 0) ? "初検" :
    (priorCountInEpisode === 1) ? "再検" : "後療";

  return { episodeStartDate: episodeStartDate, kubun: kubun, priorCountInEpisode: priorCountInEpisode };
}

function findLatestCaseRowDateInEpisode_(caseSh, caseMap, patientId, treatDate, caseNo) {
  var dates = getPatientVisitDatesFromCasesByCase_(caseSh, caseMap, patientId, caseNo);
  if (!dates.length) return null;

  var prevDates = dates.filter(function(d) { return d.getTime() < treatDate.getTime(); }).sort(function(a,b){ return a.getTime()-b.getTime(); });
  if (!prevDates.length) return null;

  var lastDate = prevDates[prevDates.length - 1];
  var gap = daysBetween_(lastDate, treatDate);
  if (gap > 30) return null;

  return lastDate;
}

function isCaseClosedAsOf_(caseRowObj, treatDate) {
  if (!caseRowObj) return false;

  var has1 = partExists_(caseRowObj, 1);
  var has2 = partExists_(caseRowObj, 2);

  var e1 = has1 ? isEnded_(caseRowObj.end1, treatDate) : true;
  var e2 = has2 ? isEnded_(caseRowObj.end2, treatDate) : true;

  return e1 && e2;
}

function buildEpisodeDatesBackwards_StopAtClosed_(caseSh, caseMap, patientId, caseNo, prevDatesAsc, currentDate, maxGap) {
  var episode = [];
  var pivot = currentDate;

  for (var i = prevDatesAsc.length - 1; i >= 0; i--) {
    var d = prevDatesAsc[i];
    var gap = daysBetween_(d, pivot);
    if (gap > maxGap) break;

    var row = findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, d, caseNo);
    if (isCaseClosedAsOf_(row, pivot)) {
      break;
    }

    episode.unshift(d);
    pivot = d;
  }
  return episode;
}

function partExists_(src, idx) {
  if (!src) return false;
  if (idx === 1) return !!(String(src.p1||"").trim() || String(src.d1||"").trim() || (src.inj1 instanceof Date));
  return !!(String(src.p2||"").trim() || String(src.d2||"").trim() || (src.inj2 instanceof Date));
}
