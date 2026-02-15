/****************************************************
 * 柔整 Ver3 金額フェーズ（最終ヘッダー前提・衝突ゼロ版）
 * - core側の SHEETS / HEADER_COLS / MASTER_COLS / buildHeaderColMap_ 等を利用する
 * - このファイルでは「新規関数」と「AM_定数」だけを持つ（重複宣言しない）
 *
 * ★最終方針
 * - 施術明細の金額は “_確定” 列のみを正とする
 * - 旧列（施療/後療料, 冷罨法, 温罨法, 電療, 明細小計 等）は参照/更新しない
 ****************************************************/

/** ===== 設定キー（設定シートA列）===== */
const AM_SET_KEYS = {
  initFee: "初検料",
  initSupport: "初検時相談支援",
  reFee: "再検料",
  shoryoDaboku: "施療料_打撲",
  shoryoNenZa: "施療料_捻挫",
  shoryoZasyo: "施療料_挫傷",
  koryoDaboku: "後療料_打撲",
  koryoNenZa: "後療料_捻挫",
  koryoZasyo: "後療料_挫傷",
  cold: "冷罨法",
  warm: "温罨法",
  electro: "電療",
  taiki: "待機_打撲捻挫挫傷",
  multiCoef3: "多部位_3部位",
  roundUnit: "窓口端数単位",
};

/** ===== 施術明細：列名（最終ヘッダー前提）===== */
const AM_DETAIL_COLS = {
  // key
  visitKey: "visitKey",
  patientId: "患者ID",
  treatDate: "施術日",
  kubun: "区分", // 初検/再検/後療
  injuryDateFixed: "受傷日_確定",
  injuryDateInput: "受傷日(入力)",
  byomei: "傷病",
  partOrder: "部位順位", // 1,2,3...
  coldChk: "冷",
  warmChk: "温",
  electroChk: "電",

  // 出力（確定列）
  coefOut: "係数",
  baseOut: "基本料_確定",
  supportOut: "初検相談_確定",
  coldOut: "冷_確定",
  warmOut: "温_確定",
  electroOut: "電_確定",
  taikiOut: "待機_確定",
  rowTotalOut: "行合計_確定",
};

/** メニュー実行（visitKeyを入力） */
function menuRecalcAmounts_V3() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const res = ui.prompt("金額再計算", "visitKey を入力してください（例：P0001|2026-02-15）", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const visitKey = (res.getResponseText() || "").trim();
  if (!visitKey) {
    ui.alert("visitKey が空です。");
    return;
  }

  const result = recalcAmountsByVisitKey_V3_(ss, visitKey);
  ui.alert(
    `完了：${visitKey}\n` +
    `明細行更新：${result.updatedRows}\n` +
    `来院合計：${result.total}\n` +
    `窓口負担：${result.copay}\n` +
    `保険請求：${result.claim}`
  );
}

/** visitKey単位で、明細→ヘッダを再計算 */
function recalcAmountsByVisitKey_V3_(ss, visitKey) {
  const settings = loadSettings_V3_(ss);

  const detailSh = ss.getSheetByName(SHEETS.detail);
  const headerSh = ss.getSheetByName(SHEETS.header);
  const masterSh = ss.getSheetByName(SHEETS.master);

  const maps = {
    detail: buildHeaderColMap_(detailSh),
    header: buildHeaderColMap_(headerSh),
    master: buildHeaderColMap_(masterSh),
  };

  // 必須列チェック（不足なら即エラー＝事故ゼロ）
  ensureRequiredCols_(maps.detail, Object.values(AM_DETAIL_COLS));
  ensureRequiredCols_(maps.header, [
    HEADER_COLS.visitKey,
    HEADER_COLS.visitTotal,
    HEADER_COLS.windowPay,
    HEADER_COLS.claimPay,
  ]);
  ensureRequiredCols_(maps.master, [MASTER_COLS.patientId, MASTER_COLS.burden]);

  // 明細全取得
  const detailValues = detailSh.getDataRange().getValues();
  if (detailValues.length < 2) throw new Error("施術明細にデータがありません。");

  // visitKey該当行を収集（0-based index）
  const vkCol0 = maps.detail[AM_DETAIL_COLS.visitKey] - 1;
  const rows0 = [];
  for (let r0 = 1; r0 < detailValues.length; r0++) {
    if (String(detailValues[r0][vkCol0] || "").trim() === visitKey) rows0.push(r0);
  }
  if (!rows0.length) throw new Error(`施術明細で visitKey=${visitKey} が見つかりません。`);

  // 患者ID
  const pidCol0 = maps.detail[AM_DETAIL_COLS.patientId] - 1;
  const patientId = String(detailValues[rows0[0]][pidCol0] || "").trim();
  if (!patientId) throw new Error("施術明細の患者IDが空です。");

  // 負担割合
  const burden = loadBurdenRatio_V3_(masterSh, maps.master, patientId);

  let total = 0;

  // 明細：あなたの設計優先（1行ずつ setValue）
  for (const r0 of rows0) {
    const row = detailValues[r0];

    const kubun = String(row[maps.detail[AM_DETAIL_COLS.kubun] - 1] || "").trim();
    const byomei = String(row[maps.detail[AM_DETAIL_COLS.byomei] - 1] || "").trim();

    const treatDate = asDate_V3_(row[maps.detail[AM_DETAIL_COLS.treatDate] - 1]);
    const injuryDate = pickInjuryDate_V3_(row, maps.detail);

    const partOrder = Number(row[maps.detail[AM_DETAIL_COLS.partOrder] - 1] || 0) || 0;
    const coef = (partOrder >= 3) ? Number(settings.multiCoef3 || 0.6) : 1.0;

    const coldChk = row[maps.detail[AM_DETAIL_COLS.coldChk] - 1] === true;
    const warmChk = row[maps.detail[AM_DETAIL_COLS.warmChk] - 1] === true;
    const electroChk = row[maps.detail[AM_DETAIL_COLS.electroChk] - 1] === true;

    const injuryType = detectInjuryType_V3_(byomei);
    const base = calcBaseFee_V3_(settings, kubun, injuryType);

    // 相談支援：運用ON列が無いので事故防止で0固定（必要なら後でON列追加）
    const support = 0;

    const dayDiff = diffDays_V3_(injuryDate, treatDate);

    // ★あなたのルール（厚労省運用）
    const cold = (coldChk && kubun === "初検" && dayDiff != null && dayDiff <= 1) ? settings.cold : 0;
    const warm = (warmChk && (kubun === "再検" || kubun === "後療") && dayDiff != null && dayDiff >= 5) ? settings.warm : 0;
    const electro = (electroChk && (kubun === "再検" || kubun === "後療") && dayDiff != null && dayDiff >= 5) ? settings.electro : 0;
    const taiki = ((warm > 0 || electro > 0) && (kubun === "再検" || kubun === "後療") && dayDiff != null && dayDiff >= 5) ? settings.taiki : 0;

    const rowTotal = (base + support + cold + warm + electro + taiki) * coef;
    total += rowTotal;

    // 書き込み（1-based行/列）
    const row1 = r0 + 1;
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.coefOut]).setValue(coef);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.baseOut]).setValue(base);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.supportOut]).setValue(support);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.coldOut]).setValue(cold);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.warmOut]).setValue(warm);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.electroOut]).setValue(electro);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.taikiOut]).setValue(taiki);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.rowTotalOut]).setValue(rowTotal);
  }

  // 窓口・請求
  const unit = settings.roundUnit || 1;
  const copayRaw = total * burden;
  const copay = roundToUnit_V3_(copayRaw, unit);
  const claim = total - copay;

  // ヘッダ行を探して更新
  const headerValues = headerSh.getDataRange().getValues();
  const hkCol0 = maps.header[HEADER_COLS.visitKey] - 1;
  let headerRow0 = -1;
  for (let r0 = 1; r0 < headerValues.length; r0++) {
    if (String(headerValues[r0][hkCol0] || "").trim() === visitKey) { headerRow0 = r0; break; }
  }
  if (headerRow0 === -1) throw new Error(`来院ヘッダで visitKey=${visitKey} が見つかりません。`);

  const hr1 = headerRow0 + 1;
  headerSh.getRange(hr1, maps.header[HEADER_COLS.visitTotal]).setValue(total);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.windowPay]).setValue(copay);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.claimPay]).setValue(claim);

  return { updatedRows: rows0.length, total, copay, claim };
}

/** 設定読み込み（A:キー B:値） */
function loadSettings_V3_(ss) {
  const sh = ss.getSheetByName(SHEETS.settings);
  const values = sh.getDataRange().getValues();
  const map = {};

  for (let r = 1; r < values.length; r++) {
    const key = String(values[r][0] || "").trim();
    if (!key) continue;
    const val = values[r][1];
    map[key] = (typeof val === "number") ? val : Number(String(val || "").trim() || 0);
  }

  return {
    initFee: Number(map[AM_SET_KEYS.initFee] || 0),
    initSupport: Number(map[AM_SET_KEYS.initSupport] || 0),
    reFee: Number(map[AM_SET_KEYS.reFee] || 0),
    shoryoDaboku: Number(map[AM_SET_KEYS.shoryoDaboku] || 0),
    shoryoNenZa: Number(map[AM_SET_KEYS.shoryoNenZa] || 0),
    shoryoZasyo: Number(map[AM_SET_KEYS.shoryoZasyo] || 0),
    koryoDaboku: Number(map[AM_SET_KEYS.koryoDaboku] || 0),
    koryoNenZa: Number(map[AM_SET_KEYS.koryoNenZa] || 0),
    koryoZasyo: Number(map[AM_SET_KEYS.koryoZasyo] || 0),
    cold: Number(map[AM_SET_KEYS.cold] || 0),
    warm: Number(map[AM_SET_KEYS.warm] || 0),
    electro: Number(map[AM_SET_KEYS.electro] || 0),
    taiki: Number(map[AM_SET_KEYS.taiki] || 0),
    multiCoef3: Number(map[AM_SET_KEYS.multiCoef3] || 0.6),
    roundUnit: Number(map[AM_SET_KEYS.roundUnit] || 1),
  };
}

/** 患者マスタから負担割合取得（0.3 or 30 どちらでもOK） */
function loadBurdenRatio_V3_(masterSh, masterMap, patientId) {
  const values = masterSh.getDataRange().getValues();
  const pidCol0 = masterMap[MASTER_COLS.patientId] - 1;
  const bCol0 = masterMap[MASTER_COLS.burden] - 1;

  for (let r0 = 1; r0 < values.length; r0++) {
    if (String(values[r0][pidCol0] || "").trim() !== patientId) continue;
    const raw = values[r0][bCol0];
    const num = (typeof raw === "number") ? raw : Number(String(raw || "").trim());
    if (!isFinite(num)) return 0;
    return (num > 1) ? (num / 100) : num;
  }
  throw new Error(`患者マスタに患者ID=${patientId}が見つかりません。`);
}

/** 基本料（あなたの設定シート通り） */
function calcBaseFee_V3_(settings, kubun, injuryType) {
  if (kubun === "初検") {
    if (injuryType === "打撲") return settings.shoryoDaboku;
    if (injuryType === "捻挫") return settings.shoryoNenZa;
    if (injuryType === "挫傷") return settings.shoryoZasyo;
    return 0;
  }
  if (kubun === "再検" || kubun === "後療") {
    if (injuryType === "打撲") return settings.koryoDaboku;
    if (injuryType === "捻挫") return settings.koryoNenZa;
    if (injuryType === "挫傷") return settings.koryoZasyo;
    return 0;
  }
  return 0;
}

function detectInjuryType_V3_(byomei) {
  if (!byomei) return null;
  if (byomei.indexOf("打撲") !== -1) return "打撲";
  if (byomei.indexOf("捻挫") !== -1) return "捻挫";
  if (byomei.indexOf("挫傷") !== -1) return "挫傷";
  return null;
}

function pickInjuryDate_V3_(row, detailMap) {
  const cFixed = detailMap[AM_DETAIL_COLS.injuryDateFixed];
  if (cFixed) {
    const d = asDate_V3_(row[cFixed - 1]);
    if (d) return d;
  }
  const cIn = detailMap[AM_DETAIL_COLS.injuryDateInput];
  if (cIn) {
    const d = asDate_V3_(row[cIn - 1]);
    if (d) return d;
  }
  return null;
}

function diffDays_V3_(injury, treat) {
  if (!(injury instanceof Date) || !(treat instanceof Date)) return null;
  const a = new Date(injury.getFullYear(), injury.getMonth(), injury.getDate());
  const b = new Date(treat.getFullYear(), treat.getMonth(), treat.getDate());
  return Math.round((b.getTime() - a.getTime()) / (24 * 3600 * 1000));
}

function asDate_V3_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

/** 端数処理（四捨五入） */
function roundToUnit_V3_(value, unit) {
  const u = Number(unit || 1);
  if (!isFinite(value)) return 0;
  if (!isFinite(u) || u <= 0) return Math.round(value);
  return Math.round(value / u) * u;
}