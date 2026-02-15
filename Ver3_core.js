/****************************************************
 * 柔整 Ver3.1（統合・1ファイル版：安全弁＋施術開始/終了つき）
 *
 * 追加：
 * A) 保存後、来院日(B4)以外をクリア
 * B) 経過履歴：前5回（参照用）
 * C) 2回目以降の自動引継ぎ（30日以内）
 * D) 部位ごとの 施術開始日/施術終了日 を来院ケースで管理し、
 *    終了した部位は「コア」を自動引継ぎしない
 *
 * ★重要（仕様）
 * - 終了日セルに「何か入っていれば」終了扱い（Dateでも文字でも）
 * - 同日も終了扱い（<=）
 * - ただし「終了でも治療はする」ため、チェックは“消さない”
 *   → 終了は「コア補完を止める」だけ
 *
 * ★追加仕様
 * - 終了ケース（両部位が終了扱い）
 *   → 次回来院時、そのcaseは
 *      ・部位/傷病/受傷日/治療チェック：自動引継ぎしない
 *      ・区分表示：初検（＝新規事象扱い）
 *      ・所見：引継ぎしない（空白維持）
 *      ・経過履歴（過去5回）：残す（参照用）
 *
 * 【安全弁】
 * 1) 保存判定は「コア（部位/傷病/受傷日）」のみ（チェックだけTRUEでは保存しない）
 * 2) 温電TRUEは「コアがある行」だけ
 * 3) コアがない行はチェックを必ずFALSEへ戻す（事故ゼロ）
 *
 * ★UI改修（2026-02-13）
 * - G列に「終了日」入力欄を追加（Case1/Case2）
 * - 所見/経過/履歴は1列右へ（G→H）
 *
 * ★B完成版（終了境界で初検リセット）
 * - エピソード連結（後ろ向き30日）を
 *   「途中に“両部位終了”があれば、そこで打ち切り」に変更
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
  // UIが取れないコンテキストでも落ちないよう安全化
  try {
    SpreadsheetApp.getUi()
      .createMenu("柔整ツール")
      .addItem("来院ケースへ保存（case1+case2）", "saveVisitToCases_V3")
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

    const sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.ui) return;

    const a1 = e.range.getA1Notation();
    if (a1 !== UI.patientId && a1 !== UI.treatDate) return;

    // ===== 同時実行/連打対策（くるくる防止）=====
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(500)) return; // 0.5秒で取れなければ捨てる（連打吸収）

    try {
      // 直近実行から短時間ならスキップ（デバウンス）
      const props = PropertiesService.getScriptProperties();
      const now = Date.now();
      const last = Number(props.getProperty("V3_ONEDIT_LAST") || 0);
      if (now - last < 1200) return; // 1.2秒以内の連続は無視
      props.setProperty("V3_ONEDIT_LAST", String(now));

      // ★ここだけ実行（重い処理はここに集約）
      refreshKeikaHistoryUI_V3();
      autofillFromPreviousVisit_V3();
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
  }
}

/** ===== 1行目ヘッダー名→列番号 ===== */
function buildHeaderColMap_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });
  return map;
}

/** ===== 結合セル（左上） ===== */
function getMergedValue_(sheet, a1Range) {
  return sheet.getRange(a1Range).getCell(1, 1).getValue();
}
function setMergedValue_(sheet, a1Range, value) {
  const cell = sheet.getRange(a1Range).getCell(1, 1);
  cell.setValue(value);
  cell.setWrap(true);
}

/** ===== 日付ユーティリティ ===== */
function fmt_(d, pat) {
  return Utilities.formatDate(d, "Asia/Tokyo", pat);
}
function daysBetween_(fromDate, toDate) {
  const a = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  const b = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
  return Math.round((b.getTime() - a.getTime()) / (24 * 3600 * 1000));
}
function minDate_(d1, d2) {
  const a = (d1 instanceof Date) ? d1 : null;
  const b = (d2 instanceof Date) ? d2 : null;
  if (!a && !b) return "";
  if (a && !b) return a;
  if (!a && b) return b;
  return (a.getTime() <= b.getTime()) ? a : b;
}

/** ===== 必須列チェック ===== */
function ensureRequiredCols_(map, requiredList) {
  const missing = requiredList.filter(n => !map[n]);
  if (missing.length) throw new Error("ヘッダー不足：\n- " + missing.join("\n- "));
}

/** ===== キー列で行検索 ===== */
function findRowByKey_(sheet, map, keyHeaderName, keyValue) {
  const c = map[keyHeaderName];
  if (!c) throw new Error(`キー列がありません: ${keyHeaderName}`);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const vals = sheet.getRange(2, c, lastRow - 1, 1).getValues().flat();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i] || "").trim() === String(keyValue)) return i + 2;
  }
  return 0;
}

/** ===== rowArrへ列名でセット ===== */
function setByName_(rowArr, headerMap, name, value, opt = {}) {
  const col = headerMap[name];
  if (!col) throw new Error(`対象シートに列がありません: ${name}`);
  const idx = col - 1;
  if (opt.preserveIfExists && rowArr[idx] !== "" && rowArr[idx] != null) return;
  rowArr[idx] = value;
}

/** ===== チェック安全弁 ===== */
function coreHasAny_(rowVals /* [部位,傷病,受傷日,冷,温,電,(終了日)] */) {
  const part = String(rowVals[0] || "").trim();
  const dis  = String(rowVals[1] || "").trim();
  const inj  = rowVals[2] instanceof Date;
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

/** ===== 終了判定（終了セルに何か入っていれば終了。同日も終了扱い） ===== */
function isEnded_(endVal, treatDate) {
  if (endVal === "" || endVal == null) return false;     // 未入力＝未終了
  if (!(endVal instanceof Date)) return true;            // 文字でも何でも入ってれば終了（事故ゼロ）
  if (!(treatDate instanceof Date)) return true;         // 来院日が不正なら安全側（終了）
  return endVal.getTime() <= treatDate.getTime();        // 同日も終了
}

/** ===== UIの2行入力を読む（保存判定はコアのみ） ===== */
function readRowNewUI_(uiSh, a1) {
  const v = uiSh.getRange(a1).getValues()[0];

  const part = String(v[0] || "").trim();
  const disease = String(v[1] || "").trim();
  const injuryDate = (v[2] instanceof Date) ? v[2] : "";

  const cold = v[3] === true;
  const warm = v[4] === true;
  const elec = v[5] === true;

  // 終了日（G列）
  const endRaw = v[6];
  const endVal =
    (endRaw instanceof Date) ? endRaw :
    (String(endRaw || "").trim() ? String(endRaw) : "");

  const hasCore = !!(part || disease || injuryDate);

  return { part, disease, injuryDate, cold, warm, elec, endVal, hasCore };
}

/** ===== 来院ケースから患者の来院日一覧（ユニーク） ===== */
function getPatientVisitDatesFromCases_(caseSh, caseMap, patientId) {
  const lastRow = caseSh.getLastRow();
  if (lastRow < 2) return [];

  const n = lastRow - 1;
  const cPid = caseMap[CASE_COLS.patientId];
  const cDt  = caseMap[CASE_COLS.treatDate];

  const pidVals = caseSh.getRange(2, cPid, n, 1).getValues().flat();
  const dtVals  = caseSh.getRange(2, cDt,  n, 1).getValues().flat();

  const set = new Map();
  for (let i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    const d = dtVals[i];
    if (!(d instanceof Date)) continue;
    const k = fmt_(d, "yyyy-MM-dd");
    if (!set.has(k)) set.set(k, d);
  }
  return Array.from(set.values()).sort((a,b)=>a.getTime()-b.getTime());
}

/** ===== 保存 ===== */
function saveVisitToCases_V3() {
  const ss = SpreadsheetApp.getActive();
  const uiSh = ss.getSheetByName(SHEETS.ui);
  const caseSh = ss.getSheetByName(SHEETS.cases);
  if (!uiSh || !caseSh) throw new Error("必要なシートが見つかりません（患者画面/来院ケース）");

  const patientId = String(uiSh.getRange(UI.patientId).getValue() || "").trim();
  const treatDate = uiSh.getRange(UI.treatDate).getValue();

  if (!patientId) throw new Error("患者ID（B2）が空です。");
  if (!(treatDate instanceof Date)) throw new Error("来院日（B4）が日付になっていません。");

  const visitKey = `${patientId}|${fmt_(treatDate, "yyyy-MM-dd")}`;
  const now = new Date();

  const caseMap = buildHeaderColMap_(caseSh);
  ensureRequiredCols_(caseMap, Object.values(CASE_COLS));

  // case単位でエピソードと区分を計算（★B完成版：終了境界で切る）
  const ep1 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 1);
  const ep2 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 2);

  upsertOneCase_(uiSh, caseSh, caseMap, {
    visitKey, patientId, treatDate,
    kubun: ep1.kubun,
    caseNo: 1,
    now,
    episodeStartDate: ep1.episodeStartDate
  });

  upsertOneCase_(uiSh, caseSh, caseMap, {
    visitKey, patientId, treatDate,
    kubun: ep2.kubun,
    caseNo: 2,
    now,
    episodeStartDate: ep2.episodeStartDate
  });

  refreshKeikaHistoryUI_V3();
  clearAfterSaveUI_V3_(uiSh);
  SpreadsheetApp.getUi().alert("来院ケースへ保存しました（case1 + case2）");
}

function upsertOneCase_(uiSh, caseSh, caseMap, base) {
  const { visitKey, patientId, treatDate, kubun, caseNo, now, episodeStartDate } = base;

  const rows = (caseNo === 1) ? UI.case1_rows : UI.case2_rows;
  const shokenRange = (caseNo === 1) ? UI.case1_shoken : UI.case2_shoken;
  const keikaRange  = (caseNo === 1) ? UI.case1_keikaNow : UI.case2_keikaNow;

  const line1 = readRowNewUI_(uiSh, rows[0]);
  const line2 = readRowNewUI_(uiSh, rows[1]);

  const shoken = String(getMergedValue_(uiSh, shokenRange) || "").trim();
  const keikaNow = String(getMergedValue_(uiSh, keikaRange) || "").trim();

  const hasAny = line1.hasCore || line2.hasCore || !!shoken || !!keikaNow || !!line1.endVal || !!line2.endVal;
  if (!hasAny) return;

  const injuryFixed = minDate_(line1.injuryDate, line2.injuryDate);
  const caseKey = `${visitKey}|C${caseNo}`;
  const rowIndex = findRowByKey_(caseSh, caseMap, CASE_COLS.caseKey, caseKey);

  if (rowIndex === 0) {
    const rowArr = new Array(caseSh.getLastColumn()).fill("");

    setByName_(rowArr, caseMap, CASE_COLS.visitKey, visitKey);
    setByName_(rowArr, caseMap, CASE_COLS.treatDate, treatDate);
    setByName_(rowArr, caseMap, CASE_COLS.patientId, patientId);
    setByName_(rowArr, caseMap, CASE_COLS.caseNo, caseNo);
    setByName_(rowArr, caseMap, CASE_COLS.caseKey, caseKey);
    setByName_(rowArr, caseMap, CASE_COLS.kubun, kubun);

    if (injuryFixed) setByName_(rowArr, caseMap, CASE_COLS.injuryFixed, injuryFixed);

    writeLinesToCaseRow_(rowArr, caseMap, line1, line2);

    // 施術開始日：初検は treatDate、初検以外は episodeStartDate
    if (line1.hasCore) setByName_(rowArr, caseMap, CASE_COLS.start1, (kubun === "初検") ? treatDate : episodeStartDate);
    if (line2.hasCore) setByName_(rowArr, caseMap, CASE_COLS.start2, (kubun === "初検") ? treatDate : episodeStartDate);

    // 施術終了日：UIの終了日（G列）
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
    const lastCol = caseSh.getLastColumn();
    const rowArr = caseSh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

    setByName_(rowArr, caseMap, CASE_COLS.visitKey, visitKey);
    setByName_(rowArr, caseMap, CASE_COLS.treatDate, treatDate);
    setByName_(rowArr, caseMap, CASE_COLS.patientId, patientId);
    setByName_(rowArr, caseMap, CASE_COLS.caseNo, caseNo);
    setByName_(rowArr, caseMap, CASE_COLS.caseKey, caseKey);
    setByName_(rowArr, caseMap, CASE_COLS.kubun, kubun);

    if (injuryFixed) setByName_(rowArr, caseMap, CASE_COLS.injuryFixed, injuryFixed, { preserveIfExists: true });

    writeLinesToCaseRow_(rowArr, caseMap, line1, line2);

    // 施術開始日：初検は treatDate で上書き、初検以外は初回値を保持
    if (line1.hasCore) {
      if (kubun === "初検") {
        setByName_(rowArr, caseMap, CASE_COLS.start1, treatDate);
      } else {
        setByName_(rowArr, caseMap, CASE_COLS.start1, episodeStartDate, { preserveIfExists: true });
      }
    }
    if (line2.hasCore) {
      if (kubun === "初検") {
        setByName_(rowArr, caseMap, CASE_COLS.start2, treatDate);
      } else {
        setByName_(rowArr, caseMap, CASE_COLS.start2, episodeStartDate, { preserveIfExists: true });
      }
    }

    // 施術終了日：UIに入力があれば反映（空なら上書きしない）
    if (line1.endVal !== "" && line1.endVal != null) setByName_(rowArr, caseMap, CASE_COLS.end1, line1.endVal);
    if (line2.endVal !== "" && line2.endVal != null) setByName_(rowArr, caseMap, CASE_COLS.end2, line2.endVal);

    if (shoken) setByName_(rowArr, caseMap, CASE_COLS.shoken, shoken);
    if (keikaNow) setByName_(rowArr, caseMap, CASE_COLS.keikaNow, keikaNow);

    caseSh.getRange(rowIndex, 1, 1, lastCol).setValues([rowArr]);
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

/** ===== 経過履歴（最新5件：参照用なので終了後も残す） ===== */
function refreshKeikaHistoryUI_V3() {
  const ss = SpreadsheetApp.getActive();
  const uiSh = ss.getSheetByName(SHEETS.ui);
  const caseSh = ss.getSheetByName(SHEETS.cases);
  if (!uiSh || !caseSh) return;

  const patientId = String(uiSh.getRange(UI.patientId).getValue() || "").trim();
  if (!patientId) {
    setMergedValue_(uiSh, UI.case1_keikaHistory, "");
    setMergedValue_(uiSh, UI.case2_keikaHistory, "");
    return;
  }

  const caseMap = buildHeaderColMap_(caseSh);
  ensureRequiredCols_(caseMap, [CASE_COLS.patientId, CASE_COLS.caseNo, CASE_COLS.treatDate, CASE_COLS.keikaNow]);

  const hist1 = buildKeikaHistoryTextFromCases_(caseSh, caseMap, patientId, 1, 5);
  const hist2 = buildKeikaHistoryTextFromCases_(caseSh, caseMap, patientId, 2, 5);

  setMergedValue_(uiSh, UI.case1_keikaHistory, hist1);
  setMergedValue_(uiSh, UI.case2_keikaHistory, hist2);
}

function buildKeikaHistoryTextFromCases_(caseSh, caseMap, patientId, caseNo, limit) {
  const cPid = caseMap[CASE_COLS.patientId];
  const cNo  = caseMap[CASE_COLS.caseNo];
  const cDt  = caseMap[CASE_COLS.treatDate];
  const cK   = caseMap[CASE_COLS.keikaNow];

  const lastRow = caseSh.getLastRow();
  if (lastRow < 2) return "";

  const pidVals = caseSh.getRange(2, cPid, lastRow - 1, 1).getValues().flat();
  const noVals  = caseSh.getRange(2, cNo,  lastRow - 1, 1).getValues().flat();
  const dtVals  = caseSh.getRange(2, cDt,  lastRow - 1, 1).getValues().flat();
  const kVals   = caseSh.getRange(2, cK,   lastRow - 1, 1).getValues().flat();

  const rows = [];
  for (let i = 0; i < pidVals.length; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    if (Number(noVals[i] || 0) !== caseNo) continue;

    const d = dtVals[i];
    const k = String(kVals[i] || "").trim();
    if (!(d instanceof Date)) continue;
    if (!k) continue;
    rows.push({ d, k });
  }

  rows.sort((a, b) => b.d.getTime() - a.d.getTime());
  return rows.slice(0, limit).map(x => `${fmt_(x.d, "M/d")}：${x.k}`).join("\n");
}

/** ===== 自動引継ぎ（30日以内のみ） ===== */
function autofillFromPreviousVisit_V3() {
  const ss = SpreadsheetApp.getActive();
  const uiSh = ss.getSheetByName(SHEETS.ui);
  const caseSh = ss.getSheetByName(SHEETS.cases);
  if (!uiSh || !caseSh) return;

  const patientId = String(uiSh.getRange(UI.patientId).getValue() || "").trim();
  const treatDate = uiSh.getRange(UI.treatDate).getValue();

  if (!patientId || !(treatDate instanceof Date)) {
    uiSh.getRange(UI.case1_kubunView).setValue("");
    uiSh.getRange(UI.case2_kubunView).setValue("");
    return;
  }

  const caseMap = buildHeaderColMap_(caseSh);
  ensureRequiredCols_(caseMap, [
    CASE_COLS.patientId, CASE_COLS.treatDate, CASE_COLS.caseNo,
    CASE_COLS.p1, CASE_COLS.d1, CASE_COLS.inj1,
    CASE_COLS.p2, CASE_COLS.d2, CASE_COLS.inj2,
    CASE_COLS.start1, CASE_COLS.end1,
    CASE_COLS.start2, CASE_COLS.end2,
    CASE_COLS.shoken
  ]);

  const ep1 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 1);
  const ep2 = calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, 2);

  const latest1 = findLatestCaseRowDateInEpisode_(caseSh, caseMap, patientId, treatDate, 1);
  const latest2 = findLatestCaseRowDateInEpisode_(caseSh, caseMap, patientId, treatDate, 2);

  const src1 = latest1 ? findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, latest1, 1) : null;
  const src2 = latest2 ? findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, latest2, 2) : null;

  /* =========================
     ===== case1 =====
     ========================= */
  {
    const caseClosed = isCaseClosedAsOf_(src1, treatDate);

    if (caseClosed) {
      // ★終了ケース＝新規事象扱い
      uiSh.getRange(UI.case1_kubunView).setValue("初検");
      setMergedValue_(uiSh, UI.case1_shoken, "");

      // コア補完しない（安全弁のみ）
      applyCaseRowToUI_Safe_(uiSh, null, 1, treatDate, { forceWarmElec: false });

    } else {
      uiSh.getRange(UI.case1_kubunView).setValue(ep1.kubun || "");

      applyCaseRowToUI_Safe_(uiSh, src1, 1, treatDate, { forceWarmElec: true });

      // 所見補完（UI空のときのみ）
      const curShoken = String(getMergedValue_(uiSh, UI.case1_shoken) || "").trim();
      if (!curShoken && src1) {
        const startForShoken = minDate_(src1.start1, src1.start2);
        const shokenDate = (startForShoken instanceof Date)
          ? startForShoken
          : ep1.episodeStartDate;

        const shokenSrc = findCaseRowByPatientDateCaseNo_(
          caseSh, caseMap, patientId, shokenDate, 1
        );

        if (shokenSrc) {
          const srcText = String(shokenSrc.shoken || "").trim();
          if (srcText) setMergedValue_(uiSh, UI.case1_shoken, srcText);
        }
      }
    }
  }

  /* =========================
     ===== case2 =====
     ========================= */
  {
    const caseClosed = isCaseClosedAsOf_(src2, treatDate);

    if (caseClosed) {
      uiSh.getRange(UI.case2_kubunView).setValue("初検");
      setMergedValue_(uiSh, UI.case2_shoken, "");

      applyCaseRowToUI_Safe_(uiSh, null, 2, treatDate, { forceWarmElec: false });

    } else {
      uiSh.getRange(UI.case2_kubunView).setValue(ep2.kubun || "");

      applyCaseRowToUI_Safe_(uiSh, src2, 2, treatDate, { forceWarmElec: true });

      const curShoken = String(getMergedValue_(uiSh, UI.case2_shoken) || "").trim();
      if (!curShoken && src2) {
        const startForShoken = minDate_(src2.start1, src2.start2);
        const shokenDate = (startForShoken instanceof Date)
          ? startForShoken
          : ep2.episodeStartDate;

        const shokenSrc = findCaseRowByPatientDateCaseNo_(
          caseSh, caseMap, patientId, shokenDate, 2
        );

        if (shokenSrc) {
          const srcText = String(shokenSrc.shoken || "").trim();
          if (srcText) setMergedValue_(uiSh, UI.case2_shoken, srcText);
        }
      }
    }
  }
}

function sameDateKey_(d) {
  return (d instanceof Date) ? fmt_(d, "yyyy-MM-dd") : "";
}

function findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, dateObj, caseNo) {
  if (!dateObj) return null;
  const lastRow = caseSh.getLastRow();
  if (lastRow < 2) return null;

  const n = lastRow - 1;
  const cPid = caseMap[CASE_COLS.patientId];
  const cDt  = caseMap[CASE_COLS.treatDate];
  const cNo  = caseMap[CASE_COLS.caseNo];

  const pidVals = caseSh.getRange(2, cPid, n, 1).getValues().flat();
  const dtVals  = caseSh.getRange(2, cDt,  n, 1).getValues().flat();
  const noVals  = caseSh.getRange(2, cNo,  n, 1).getValues().flat();

  const targetKey = sameDateKey_(dateObj);

  for (let i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    if (Number(noVals[i] || 0) !== caseNo) continue;
    const d = dtVals[i];
    if (!(d instanceof Date)) continue;
    if (sameDateKey_(d) !== targetKey) continue;

    const rowIndex = i + 2;
    const row = caseSh.getRange(rowIndex, 1, 1, caseSh.getLastColumn()).getValues()[0];
    const get = (name) => row[caseMap[name] - 1];

    return {
      p1: String(get(CASE_COLS.p1) || ""),
      d1: String(get(CASE_COLS.d1) || ""),
      inj1: (get(CASE_COLS.inj1) instanceof Date) ? get(CASE_COLS.inj1) : "",
      cold1: get(CASE_COLS.cold1) === true,
      warm1: get(CASE_COLS.warm1) === true,
      elec1: get(CASE_COLS.elec1) === true,
      start1: (get(CASE_COLS.start1) instanceof Date) ? get(CASE_COLS.start1) : "",
      end1: get(CASE_COLS.end1), // 生値（文字でも拾う）

      p2: String(get(CASE_COLS.p2) || ""),
      d2: String(get(CASE_COLS.d2) || ""),
      inj2: (get(CASE_COLS.inj2) instanceof Date) ? get(CASE_COLS.inj2) : "",
      cold2: get(CASE_COLS.cold2) === true,
      warm2: get(CASE_COLS.warm2) === true,
      elec2: get(CASE_COLS.elec2) === true,
      start2: (get(CASE_COLS.start2) instanceof Date) ? get(CASE_COLS.start2) : "",
      end2: get(CASE_COLS.end2), // 生値（文字でも拾う）

      shoken: String(get(CASE_COLS.shoken) || ""),
    };
  }
  return null;
}

/** ===== UIへ安全に適用（終了部位はコア補完しない。チェックは消さない） ===== */
function applyCaseRowToUI_Safe_(uiSh, src, caseNo, treatDate, opt = {}) {
  const rows = (caseNo === 1) ? UI.case1_rows : UI.case2_rows;

  const rng1 = uiSh.getRange(rows[0]);
  const rng2 = uiSh.getRange(rows[1]);

  const v1 = rng1.getValues()[0]; // [部位,傷病,受傷日,冷,温,電,終了日]
  const v2 = rng2.getValues()[0];

  // srcが無い＝引継ぎ無し。ただし安全弁でコア無しチェックは落とす
  if (!src) {
    if (!coreHasAny_(v1)) forceChecksFalse_(v1);
    if (!coreHasAny_(v2)) forceChecksFalse_(v2);
    rng1.setValues([v1]);
    rng2.setValues([v2]);
    return;
  }

  const ended1 = isEnded_(src.end1, treatDate);
  const ended2 = isEnded_(src.end2, treatDate);

  // ===== 終了日も引用（UIのG列が空のときだけ）=====
  const hasSrc1 = !!(String(src.p1||"").trim() || String(src.d1||"").trim() || (src.inj1 instanceof Date));
  const hasSrc2 = !!(String(src.p2||"").trim() || String(src.d2||"").trim() || (src.inj2 instanceof Date));

  // v1[6], v2[6] は UIの終了日セル（G列）
  if (hasSrc1 && (v1[6] === "" || v1[6] == null)) {
    // src.end1 は Dateでも文字でも入れる（生値をそのまま）
    if (src.end1 !== "" && src.end1 != null) v1[6] = src.end1;
  }
  if (hasSrc2 && (v2[6] === "" || v2[6] == null)) {
    if (src.end2 !== "" && src.end2 != null) v2[6] = src.end2;
  }


  // 終了でもチェックは消さない（治療はする）
  // 終了は「コアの自動補完を止める」だけ
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

  const row1HasCore = coreHasAny_(v1);
  const row2HasCore = coreHasAny_(v2);

  // コアが無い行は必ずチェックを落とす（事故ゼロ）
  if (!row1HasCore) forceChecksFalse_(v1);
  if (!row2HasCore) forceChecksFalse_(v2);

  // 温電を強制ON（コアがある行のみ）
  if (opt.forceWarmElec) {
    if (row1HasCore) forceWarmElecTrue_(v1);
    if (row2HasCore) forceWarmElecTrue_(v2);
  }

  rng1.setValues([v1]);
  rng2.setValues([v2]);
}

/** ===== 来院ケース → 来院ヘッダへ一括出力（高速） ===== */
function exportHeaderFromCases_V3() {
  const ss = SpreadsheetApp.getActive();
  const caseSh = ss.getSheetByName(SHEETS.cases);
  const headSh = ss.getSheetByName(SHEETS.header);
  if (!caseSh || !headSh) throw new Error("必要シートが見つかりません（来院ケース/来院ヘッダ）");

  const caseMap = buildHeaderColMap_(caseSh);
  const headMap = buildHeaderColMap_(headSh);

  ensureRequiredCols_(caseMap, [CASE_COLS.visitKey, CASE_COLS.treatDate, CASE_COLS.patientId, CASE_COLS.kubun, CASE_COLS.injuryFixed, CASE_COLS.caseKey, CASE_COLS.caseNo]);
  ensureRequiredCols_(headMap, Object.values(HEADER_COLS));

  const lastRow = caseSh.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("来院ケースにデータがありません。");
    return;
  }

  const n = lastRow - 1;
  const visitKeyVals = caseSh.getRange(2, caseMap[CASE_COLS.visitKey], n, 1).getValues().flat();
  const treatVals    = caseSh.getRange(2, caseMap[CASE_COLS.treatDate], n, 1).getValues().flat();
  const pidVals      = caseSh.getRange(2, caseMap[CASE_COLS.patientId], n, 1).getValues().flat();
  const kubunVals    = caseSh.getRange(2, caseMap[CASE_COLS.kubun], n, 1).getValues().flat();
  const injFixVals   = caseSh.getRange(2, caseMap[CASE_COLS.injuryFixed], n, 1).getValues().flat();
  const caseKeyVals  = caseSh.getRange(2, caseMap[CASE_COLS.caseKey], n, 1).getValues().flat();
  const caseNoVals   = caseSh.getRange(2, caseMap[CASE_COLS.caseNo], n, 1).getValues().flat();

  const existed = buildExistingHeaderKeySet_(headSh, headMap);

  const out = [];
  const now = new Date();

  for (let i = 0; i < n; i++) {
    const visitKey = String(visitKeyVals[i] || "").trim();
    const patientId = String(pidVals[i] || "").trim();
    const treatDate = treatVals[i];
    const kubun = String(kubunVals[i] || "").trim();
    const inj = injFixVals[i];
    const caseKey = String(caseKeyVals[i] || "").trim();
    const caseIndex = Number(caseNoVals[i] || 0);

    if (!visitKey || !patientId || !(treatDate instanceof Date) || !caseKey || !caseIndex) continue;

    const uniq = `${caseKey}#${caseIndex}`;
    if (existed.has(uniq)) continue;

    const rowArr = new Array(headSh.getLastColumn()).fill("");

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

    const last = findLastVisitDateInHeader_(headSh, headMap, patientId, treatDate);
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
  SpreadsheetApp.getUi().alert(`来院ヘッダへ出力しました：${out.length} 行`);
}

function buildExistingHeaderKeySet_(headSh, headMap) {
  const set = new Set();
  const lastRow = headSh.getLastRow();
  if (lastRow < 2) return set;

  const cCaseKey = headMap[HEADER_COLS.caseKey];
  const cCaseIdx = headMap[HEADER_COLS.caseIndex];
  if (!cCaseKey || !cCaseIdx) return set;

  const keys = headSh.getRange(2, cCaseKey, lastRow - 1, 1).getValues().flat();
  const idxs = headSh.getRange(2, cCaseIdx, lastRow - 1, 1).getValues().flat();

  for (let i = 0; i < keys.length; i++) {
    const k = String(keys[i] || "").trim();
    const n = Number(idxs[i] || 0);
    if (k && n) set.add(`${k}#${n}`);
  }
  return set;
}

function findLastVisitDateInHeader_(headSh, headMap, patientId, treatDate) {
  const lastRow = headSh.getLastRow();
  if (lastRow < 2) return null;

  const cPid = headMap[HEADER_COLS.patientId];
  const cDt  = headMap[HEADER_COLS.treatDate];
  if (!cPid || !cDt) return null;

  const pidVals = headSh.getRange(2, cPid, lastRow - 1, 1).getValues().flat();
  const dtVals  = headSh.getRange(2, cDt,  lastRow - 1, 1).getValues().flat();

  let best = null;
  for (let i = 0; i < pidVals.length; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    const d = dtVals[i];
    if (!(d instanceof Date)) continue;
    if (d.getTime() >= treatDate.getTime()) continue;
    if (!best || d.getTime() > best.getTime()) best = d;
  }
  return best;
}

/** ===== クリア（入力だけ） ===== */
function clearEntryUI_V3() {
  const ss = SpreadsheetApp.getActive();
  const uiSh = ss.getSheetByName(SHEETS.ui);

  /* =========================================================
     🔹 基本情報エリア（B4=来院日は保持する）
  ========================================================== */

  // B2：患者ID（手入力）
  uiSh.getRange("B2").clearContent();

  // F2:F4：全額・窓口負担・備考（自動計算表示エリア）
  uiSh.getRange("F2:F4").clearContent();

  // B5:B7：区分・保険種別・窓口負担割合など
  uiSh.getRange("B5:B7").clearContent();


  /* =========================================================
     🔹 Case1（緑エリア）
  ========================================================== */

  // A12:G13：Case1 部位入力（最大2部位）
  // A=部位 / B=傷病 / C=受傷日 / D=冷 / E=温 / F=電 / G=終了日
  uiSh.getRange("A12:G13").clearContent();

  // H11:M14：所見 Case1（初検所見・表示エリア）
  uiSh.getRange("H11:M14").clearContent();

  // H16:M17：経過（今回）Case1
  uiSh.getRange("H16:M17").clearContent();

  // H19:M23：経過履歴 Case1（参照用表示）
  uiSh.getRange("H19:M23").clearContent();


  /* =========================================================
     🔹 Case2（黄エリア）
  ========================================================== */

  // A27:G28：Case2 部位入力（最大2部位）
  // A=部位 / B=傷病 / C=受傷日 / D=冷 / E=温 / F=電 / G=終了日
  uiSh.getRange("A27:G28").clearContent();

  // H26:M29：所見 Case2
  uiSh.getRange("H26:M29").clearContent();

  // H31:M32：経過（今回）Case2
  uiSh.getRange("H31:M32").clearContent();

  // H34:M38：経過履歴 Case2（参照用表示）
  uiSh.getRange("H34:M38").clearContent();


  /* =========================================================
     🔹 治療チェック欄を必ずFALSEへ（事故防止）
  ========================================================== */

  uiSh.getRange("D12:F13").setValues([
    [false, false, false],
    [false, false, false]
  ]);

  uiSh.getRange("D27:F28").setValues([
    [false, false, false],
    [false, false, false]
  ]);


  /* =========================================================
     🔹 表示専用：区分ビューもクリア
  ========================================================== */

  uiSh.getRange(UI.case1_kubunView).clearContent();
  uiSh.getRange(UI.case2_kubunView).clearContent();

  SpreadsheetApp.getUi().alert("自動入力エリアをクリアしました（B4は保持・書式は保持）。");
}

/** ===== 保存後クリア：来院日(B4)以外 ===== */
function clearAfterSaveUI_V3_(uiSh) {
  uiSh.getRange(UI.patientId).setValue("");
  uiSh.getRange(UI.kubun).setValue("");

  UI.case1_rows.forEach(r => uiSh.getRange(r).clearContent());
  UI.case2_rows.forEach(r => uiSh.getRange(r).clearContent());

  uiSh.getRange("D12:F13").setValues([[false,false,false],[false,false,false]]);
  uiSh.getRange("D27:F28").setValues([[false,false,false],[false,false,false]]);

  setMergedValue_(uiSh, UI.case1_shoken, "");
  setMergedValue_(uiSh, UI.case1_keikaNow, "");
  setMergedValue_(uiSh, UI.case2_shoken, "");
  setMergedValue_(uiSh, UI.case2_keikaNow, "");

  // 履歴は「参照用」だが、保存後は画面を軽くするため空でOK（運用どおり）
  setMergedValue_(uiSh, UI.case1_keikaHistory, "");
  setMergedValue_(uiSh, UI.case2_keikaHistory, "");

  uiSh.getRange(UI.case1_kubunView).setValue("");
  uiSh.getRange(UI.case2_kubunView).setValue("");
}

/** ===== ヘッダー確認 ===== */
function checkHeaders_V3() {
  const ss = SpreadsheetApp.getActive();
  const caseSh = ss.getSheetByName(SHEETS.cases);
  const headSh = ss.getSheetByName(SHEETS.header);
  if (!caseSh || !headSh) throw new Error("来院ケース または 来院ヘッダ が見つかりません。");

  const caseMap = buildHeaderColMap_(caseSh);
  const headMap = buildHeaderColMap_(headSh);

  const needCase = Object.values(CASE_COLS);
  const needHead = Object.values(HEADER_COLS);

  const missCase = needCase.filter(h => !caseMap[h]);
  const missHead = needHead.filter(h => !headMap[h]);

  SpreadsheetApp.getUi().alert(
    "ヘッダーチェック結果\n\n" +
    (missCase.length ? "【来院ケース 不足】\n- " + missCase.join("\n- ") + "\n\n" : "【来院ケース 不足】なし\n\n") +
    (missHead.length ? "【来院ヘッダ 不足】\n- " + missHead.join("\n- ") : "【来院ヘッダ 不足】なし")
  );
}

/** ===== caseNo別の来院日（コアがある日だけ） ===== */
function getPatientVisitDatesFromCasesByCase_(caseSh, caseMap, patientId, caseNo) {
  const lastRow = caseSh.getLastRow();
  if (lastRow < 2) return [];

  const n = lastRow - 1;
  const cPid = caseMap[CASE_COLS.patientId];
  const cDt  = caseMap[CASE_COLS.treatDate];
  const cNo  = caseMap[CASE_COLS.caseNo];

  const cP1 = caseMap[CASE_COLS.p1];
  const cD1 = caseMap[CASE_COLS.d1];
  const cI1 = caseMap[CASE_COLS.inj1];
  const cP2 = caseMap[CASE_COLS.p2];
  const cD2 = caseMap[CASE_COLS.d2];
  const cI2 = caseMap[CASE_COLS.inj2];

  const pidVals = caseSh.getRange(2, cPid, n, 1).getValues().flat();
  const dtVals  = caseSh.getRange(2, cDt,  n, 1).getValues().flat();
  const noVals  = caseSh.getRange(2, cNo,  n, 1).getValues().flat();

  const p1Vals = caseSh.getRange(2, cP1, n, 1).getValues().flat();
  const d1Vals = caseSh.getRange(2, cD1, n, 1).getValues().flat();
  const i1Vals = caseSh.getRange(2, cI1, n, 1).getValues().flat();
  const p2Vals = caseSh.getRange(2, cP2, n, 1).getValues().flat();
  const d2Vals = caseSh.getRange(2, cD2, n, 1).getValues().flat();
  const i2Vals = caseSh.getRange(2, cI2, n, 1).getValues().flat();

  const uniq = new Map();

  for (let i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    if (Number(noVals[i] || 0) !== caseNo) continue;

    const d = dtVals[i];
    if (!(d instanceof Date)) continue;

    const hasCore =
      String(p1Vals[i] || "").trim() ||
      String(d1Vals[i] || "").trim() ||
      (i1Vals[i] instanceof Date) ||
      String(p2Vals[i] || "").trim() ||
      String(d2Vals[i] || "").trim() ||
      (i2Vals[i] instanceof Date);

    if (!hasCore) continue;

    const key = fmt_(d, "yyyy-MM-dd");
    if (!uniq.has(key)) uniq.set(key, d);
  }

  return Array.from(uniq.values()).sort((a, b) => a.getTime() - b.getTime());
}

/** ===== caseNoごとのエピソード計算（30日ルール＋終了境界で打ち切り） ===== */
function calcEpisodeForCase_(caseSh, caseMap, patientId, treatDate, caseNo) {
  const dates = getPatientVisitDatesFromCasesByCase_(caseSh, caseMap, patientId, caseNo);

  if (!dates.length) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  const prevDates = dates
    .filter(d => d.getTime() < treatDate.getTime())
    .sort((a, b) => a.getTime() - b.getTime());

  if (!prevDates.length) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  // 直近の前回が30日超なら初検
  const lastDate = prevDates[prevDates.length - 1];
  const gap = daysBetween_(lastDate, treatDate);
  if (gap > 30) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  // ★B核心：後ろ向き連結を「終了境界で打ち切り」
  const episode = buildEpisodeDatesBackwards_StopAtClosed_(
    caseSh, caseMap, patientId, caseNo, prevDates, treatDate, 30
  );

  // 連結が0なら（直近が閉じ境界だった等）初検
  if (!episode.length) {
    return { episodeStartDate: treatDate, kubun: "初検", priorCountInEpisode: 0 };
  }

  const episodeStartDate = episode[0];
  const priorCountInEpisode = episode.length;

  const kubun =
    (priorCountInEpisode === 0) ? "初検" :
    (priorCountInEpisode === 1) ? "再検" : "後療";

  return { episodeStartDate, kubun, priorCountInEpisode };
}

/** ===== エピソード内の直近来院日（前回日） ===== */
function findLatestCaseRowDateInEpisode_(caseSh, caseMap, patientId, treatDate, caseNo) {
  const dates = getPatientVisitDatesFromCasesByCase_(caseSh, caseMap, patientId, caseNo);
  if (!dates.length) return null;

  const prevDates = dates.filter(d => d.getTime() < treatDate.getTime()).sort((a,b)=>a.getTime()-b.getTime());
  if (!prevDates.length) return null;

  const lastDate = prevDates[prevDates.length - 1];
  const gap = daysBetween_(lastDate, treatDate);
  if (gap > 30) return null;

  return lastDate;
}

function isCaseClosedAsOf_(caseRowObj, treatDate) {
  if (!caseRowObj) return false;

  const has1 = partExists_(caseRowObj, 1);
  const has2 = partExists_(caseRowObj, 2);

  const e1 = has1 ? isEnded_(caseRowObj.end1, treatDate) : true; // 存在しない部位は終了扱い
  const e2 = has2 ? isEnded_(caseRowObj.end2, treatDate) : true;

  return e1 && e2;
}

/**
 * 後ろ向き連結（30日）を「終了境界で打ち切り」にした版
 * - pivot（より新しい日）時点で、その日(d)のケースが「両部位終了」なら break
 */
function buildEpisodeDatesBackwards_StopAtClosed_(caseSh, caseMap, patientId, caseNo, prevDatesAsc, currentDate, maxGap) {
  const episode = [];
  let pivot = currentDate;

  for (let i = prevDatesAsc.length - 1; i >= 0; i--) {
    const d = prevDatesAsc[i];
    const gap = daysBetween_(d, pivot);
    if (gap > maxGap) break;

    // ★核心：この日が「両部位終了」なら、そこでエピソードを切る
    const row = findCaseRowByPatientDateCaseNo_(caseSh, caseMap, patientId, d, caseNo);
    if (isCaseClosedAsOf_(row, pivot)) {
      break;
    }

    episode.unshift(d);
    pivot = d;
  }
  return episode;
}

function partExists_(src, idx /*1 or 2*/) {
  if (!src) return false;
  if (idx === 1) return !!(String(src.p1||"").trim() || String(src.d1||"").trim() || (src.inj1 instanceof Date));
  return !!(String(src.p2||"").trim() || String(src.d2||"").trim() || (src.inj2 instanceof Date));
}