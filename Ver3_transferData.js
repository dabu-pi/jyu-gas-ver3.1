/****************************************************
 * 柔整 Ver3 申請書_転記データ（最終ヘッダー前提）
 * - 内部（施術明細/来院ケース/患者マスタ/保険者情報）から
 *   申請書フォーマットに近い「月次転記用データ」を生成（upsert）
 *
 * ★重要：Apps Scriptは全.gsでグローバル共有
 * - const SHEETS / HEADER_COLS / onOpen を再宣言しない（衝突防止）
 * - このファイルは V3TR 名前空間で完結させる
 ****************************************************/

var V3TR = V3TR || {};

/** ===== 設定（必要なら列名だけあなたの実シートに合わせて変更） ===== */
V3TR.CONFIG = {
  sheetNames: {
    settings: "設定",
    cases: "来院ケース",
    detail: "施術明細",
    master: "患者マスタ",
    insurer: "保険者情報",
    transfer: "申請書_転記データ",
  },

  setKeys: {
    initFee: "初検料",
    initSupport: "初検時相談支援",
    reFee: "再検料",
    roundUnit: "窓口端数単位",
  },

  masterCols: {
    patientId: "患者ID",
    name: "氏名",
    birthday: "生年月日",
    address: "住所",
    relation: "続柄",
    insuredName: "被保険者氏名",
    burdenRatio: "負担割合",
    burdenRatioDigit: "一部負担金割合", // DP45用（3/2/1）
  },

  insurerCols: {
    patientId: "患者ID",
    insurerNo: "保険者番号",
    symbol: "記号",
    number: "番号",
    insurerName: "保険者名",
  },

  caseCols: {
    patientId: "患者ID",
    treatDate: "施術日",
    caseNo: "caseNo",
    caseKey: "caseKey",
    kubun: "区分",

    p1: "部位_部位1",
    d1: "傷病_部位1",
    inj1: "受傷日_部位1",
    p2: "部位_部位2",
    d2: "傷病_部位2",
    inj2: "受傷日_部位2",

    start1: "施術開始日_部位1",
    start2: "施術開始日_部位2",
    end1: "施術終了日_部位1",
    end2: "施術終了日_部位2",
  },

  /** ★最終ヘッダー前提：施術明細（確定列を参照） */
  detailCols: {
    patientId: "患者ID",
    treatDate: "施術日",
    kubun: "区分",
    caseKey: "caseKey",

    baseFixed: "基本料_確定",
    coldFixed: "冷_確定",
    warmFixed: "温_確定",
    elecFixed: "電_確定",
    rowTotalFixed: "行合計_確定",
  },

  transferCols: [
    "recordKey", "患者ID", "対象月", "caseNo", "caseKey",

    "患者氏名", "患者生年月日", "住所", "続柄",
    "被保険者氏名", "保険者番号", "記号", "番号", "保険者名",
    "一部負担金割合",

    "負傷名", "負傷年月日", "初検年月日", "施術開始年月日", "施術終了年月日", "実日数",

    "初検料_月額", "初検時相談支援料_月額", "再検料_月額", "基本3項目_計",

    "後療料_単価", "後療料_回数", "後療料_計",
    "冷罨法_回数", "冷罨法_金額",
    "温罨法_回数", "温罨法_金額",
    "電療_回数", "電療_金額",
    "case計",

    "当月合計", "窓口負担額", "請求金額"
  ],
};

function V3TR_menuBuildTransferData() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const r1 = ui.prompt("申請書_転記データ作成", "患者ID を入力してください（例：P0001）", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const patientId = (r1.getResponseText() || "").trim();
  if (!patientId) return ui.alert("患者IDが空です。");

  const r2 = ui.prompt("対象月", "対象月（yyyy-MM）を入力してください（例：2026-02）", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const ym = (r2.getResponseText() || "").trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) return ui.alert("形式が違います。yyyy-MM で入力してください。");

  const out = V3TR_buildTransferDataForMonth_(ss, patientId, ym);
  ui.alert(
    `完了：${patientId} / ${ym}\n` +
    `更新（upsert）: ${out.upserted} 行\n` +
    `当月合計: ${out.total} / 窓口: ${out.copay} / 請求: ${out.claim}`
  );
}

function V3TR_buildTransferDataForMonth_(ss, patientId, ym) {
  const C = V3TR.CONFIG;

  const shSettings = ss.getSheetByName(C.sheetNames.settings);
  const shCases    = ss.getSheetByName(C.sheetNames.cases);
  const shDetail   = ss.getSheetByName(C.sheetNames.detail);
  const shMaster   = ss.getSheetByName(C.sheetNames.master);
  const shInsurer  = ss.getSheetByName(C.sheetNames.insurer);

  if (!shSettings || !shCases || !shDetail || !shMaster) {
    throw new Error("必要シートが見つかりません（設定/来院ケース/施術明細/患者マスタ）。");
  }

  const shTransfer = V3TR_ensureTransferSheet_(ss);

  const settings = V3TR_loadSettings_(shSettings);
  const master   = V3TR_loadMasterRow_(shMaster, patientId);
  const insurer  = (shInsurer) ? V3TR_loadInsurerRow_(shInsurer, patientId) : {};

  const month = V3TR_parseYM_(ym);
  const start = month.start;
  const end   = month.end; // exclusive

  const caseSummary = V3TR_buildCaseMonthlySummary_(shCases, patientId, start, end);
  const detailAgg = V3TR_aggregateDetailMonthly_(shDetail, patientId, start, end);
  const kubunCount = V3TR_countKubunInCases_(shCases, patientId, start, end);

  const initFee   = settings.initFee * kubunCount.initCount;
  const support   = settings.initSupport * kubunCount.initCount;
  const reFee     = settings.reFee * kubunCount.reCount;
  const base3sum  = initFee + support + reFee;

  const c1 = V3TR_buildCaseMoneyBlock_(detailAgg.case1);
  const c2 = V3TR_buildCaseMoneyBlock_(detailAgg.case2);

  const total = base3sum + c1.caseTotal + c2.caseTotal;

  const burdenRatio = V3TR_normRatio_(master.burdenRatio);
  const copay = V3TR_roundToUnit_(total * burdenRatio, settings.roundUnit);
  const claim = Math.floor(total - copay); // 小数点以下切り捨て

  const rowsOut = [];
  for (const caseNo of [1, 2]) {
    const key = `${patientId}|${ym}|C${caseNo}`;

    const cs = (caseNo === 1) ? caseSummary.case1 : caseSummary.case2;
    const cm = (caseNo === 1) ? c1 : c2;

    const jitsunisu = (caseNo === 1) ? detailAgg.case1.visitDays : detailAgg.case2.visitDays;

    const row = {};
    row["recordKey"] = key;
    row["患者ID"] = patientId;
    row["対象月"] = ym;
    row["caseNo"] = caseNo;
    row["caseKey"] = cs.caseKey || "";

    row["患者氏名"] = master.name || "";
    row["患者生年月日"] = master.birthday || "";
    row["住所"] = master.address || "";
    row["続柄"] = master.relation || "";

    row["被保険者氏名"] = master.insuredName || "";
    row["保険者番号"] = insurer.insurerNo || "";
    row["記号"] = insurer.symbol || "";
    row["番号"] = insurer.number || "";
    row["保険者名"] = insurer.insurerName || "";
    row["一部負担金割合"] = V3TR_pickBurdenDigit_(master) || "";

    row["負傷名"] = cs.injuryName || "";
    row["負傷年月日"] = cs.injuryDate || "";
    row["初検年月日"] = cs.firstDate || "";
    row["施術開始年月日"] = cs.startDate || "";
    row["施術終了年月日"] = cs.endDate || "";
    row["実日数"] = jitsunisu || 0;

    if (caseNo === 1) {
      row["初検料_月額"] = initFee;
      row["初検時相談支援料_月額"] = support;
      row["再検料_月額"] = reFee;
      row["基本3項目_計"] = base3sum;
    } else {
      row["初検料_月額"] = "";
      row["初検時相談支援料_月額"] = "";
      row["再検料_月額"] = "";
      row["基本3項目_計"] = "";
    }

    row["後療料_単価"] = cm.koryoUnit;
    row["後療料_回数"] = cm.koryoCount;
    row["後療料_計"] = cm.koryoSum;

    row["冷罨法_回数"] = cm.coldCount;
    row["冷罨法_金額"] = cm.coldSum;

    row["温罨法_回数"] = cm.warmCount;
    row["温罨法_金額"] = cm.warmSum;

    row["電療_回数"] = cm.elecCount;
    row["電療_金額"] = cm.elecSum;

    row["case計"] = cm.caseTotal;

    if (caseNo === 1) {
      row["当月合計"] = total;
      row["窓口負担額"] = copay;
      row["請求金額"] = claim;
    } else {
      row["当月合計"] = "";
      row["窓口負担額"] = "";
      row["請求金額"] = "";
    }

    rowsOut.push(row);
  }

  const upserted = V3TR_upsertTransferRows_(shTransfer, rowsOut);
  return { upserted, total, copay, claim };
}

function V3TR_ensureTransferSheet_(ss) {
  const C = V3TR.CONFIG;
  let sh = ss.getSheetByName(C.sheetNames.transfer);
  if (!sh) sh = ss.insertSheet(C.sheetNames.transfer);

  const headers = C.transferCols;
  const lastCol = sh.getLastColumn();
  const cur = (lastCol >= 1) ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  const needWrite = headers.some((h, i) => String(cur[i] || "").trim() !== h);
  if (needWrite) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function V3TR_loadSettings_(shSettings) {
  const C = V3TR.CONFIG;
  const v = shSettings.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < v.length; r++) {
    const k = String(v[r][0] || "").trim();
    if (!k) continue;
    map[k] = v[r][1];
  }
  return {
    initFee: Number(map[C.setKeys.initFee] || 0),
    initSupport: Number(map[C.setKeys.initSupport] || 0),
    reFee: Number(map[C.setKeys.reFee] || 0),
    roundUnit: Number(map[C.setKeys.roundUnit] || 1),
  };
}

function V3TR_buildHeaderMap_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  const h = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x || "").trim());
  const m = {};
  h.forEach((name, idx) => { if (name) m[name] = idx; });
  return m;
}
function V3TR_mustCol_(map, name, label) {
  const c = map[name];
  if (c === undefined) throw new Error(`${label} に列「${name}」がありません（ヘッダ行を確認）`);
  return c;
}

function V3TR_loadMasterRow_(shMaster, patientId) {
  const C = V3TR.CONFIG;
  const map = V3TR_buildHeaderMap_(shMaster);
  const v = shMaster.getDataRange().getValues();
  const cPid = V3TR_mustCol_(map, C.masterCols.patientId, "患者マスタ");

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;

    const get = (colName) => {
      const c = map[colName];
      return (c === undefined) ? "" : v[r][c];
    };

    return {
      name: get(C.masterCols.name),
      birthday: (get(C.masterCols.birthday) instanceof Date) ? get(C.masterCols.birthday) : "",
      address: get(C.masterCols.address),
      relation: get(C.masterCols.relation),
      insuredName: get(C.masterCols.insuredName),
      burdenRatio: get(C.masterCols.burdenRatio),
      burdenDigit: get(C.masterCols.burdenRatioDigit),
    };
  }
  throw new Error(`患者マスタに患者ID=${patientId}が見つかりません。`);
}

function V3TR_loadInsurerRow_(shInsurer, patientId) {
  const C = V3TR.CONFIG;
  const map = V3TR_buildHeaderMap_(shInsurer);
  const v = shInsurer.getDataRange().getValues();
  if (v.length < 2) return {};

  const cPid = map[C.insurerCols.patientId];
  if (cPid === undefined) return {};

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;
    const get = (colName) => {
      const c = map[colName];
      return (c === undefined) ? "" : v[r][c];
    };
    return {
      insurerNo: String(get(C.insurerCols.insurerNo) || "").trim(),
      symbol: String(get(C.insurerCols.symbol) || "").trim(),
      number: String(get(C.insurerCols.number) || "").trim(),
      insurerName: String(get(C.insurerCols.insurerName) || "").trim(),
    };
  }
  return {};
}

function V3TR_parseYM_(ym) {
  const [y, m] = ym.split("-").map(n => Number(n));
  const start = new Date(y, m - 1, 1);
  const end = new Date(y, m, 1);
  return { start, end };
}
function V3TR_inRange_(d, start, end) {
  if (!(d instanceof Date)) return false;
  return d.getTime() >= start.getTime() && d.getTime() < end.getTime();
}
function V3TR_dateKey_(d) {
  if (!(d instanceof Date)) return "";
  return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
}

/** ===== 来院ケース：月次代表（case1/case2） ===== */
function V3TR_buildCaseMonthlySummary_(shCases, patientId, start, end) {
  const C = V3TR.CONFIG;
  const map = V3TR_buildHeaderMap_(shCases);
  const v = shCases.getDataRange().getValues();
  if (v.length < 2) return { case1: {}, case2: {} };

  const col = (n) => V3TR_mustCol_(map, n, "来院ケース");

  const cPid = col(C.caseCols.patientId);
  const cDt  = col(C.caseCols.treatDate);
  const cNo  = col(C.caseCols.caseNo);
  const cKey = col(C.caseCols.caseKey);

  const cP1 = col(C.caseCols.p1), cD1 = col(C.caseCols.d1), cI1 = col(C.caseCols.inj1);
  const cP2 = col(C.caseCols.p2), cD2 = col(C.caseCols.d2), cI2 = col(C.caseCols.inj2);
  const cS1 = col(C.caseCols.start1), cS2 = col(C.caseCols.start2);
  const cE1 = col(C.caseCols.end1),   cE2 = col(C.caseCols.end2);

  const lastByCaseNo = { 1: null, 2: null };

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;
    const dt = v[r][cDt];
    if (!V3TR_inRange_(dt, start, end)) continue;
    const no = Number(v[r][cNo] || 0);
    if (no !== 1 && no !== 2) continue;

    if (!lastByCaseNo[no] || dt.getTime() > lastByCaseNo[no].dt.getTime()) {
      lastByCaseNo[no] = { dt, row: v[r] };
    }
  }

  function build(no) {
    const obj = lastByCaseNo[no];
    if (!obj) return { caseKey: "", injuryName: "", injuryDate: "", firstDate: "", startDate: "", endDate: "" };

    const row = obj.row;

    const p1 = String(row[cP1] || "").trim();
    const d1 = String(row[cD1] || "").trim();
    const p2 = String(row[cP2] || "").trim();
    const d2 = String(row[cD2] || "").trim();

    const injuryName = (p1 || d1) ? `${p1} ${d1}`.trim() : `${p2} ${d2}`.trim();

    const i1 = (row[cI1] instanceof Date) ? row[cI1] : "";
    const i2 = (row[cI2] instanceof Date) ? row[cI2] : "";
    const injuryDate = V3TR_minDate_(i1, i2);

    const s1 = (row[cS1] instanceof Date) ? row[cS1] : "";
    const s2 = (row[cS2] instanceof Date) ? row[cS2] : "";
    const startDate = V3TR_minDate_(s1, s2);

    const firstDate = startDate;

    const e1 = (row[cE1] instanceof Date) ? row[cE1] : "";
    const e2 = (row[cE2] instanceof Date) ? row[cE2] : "";
    const endDate = V3TR_maxDate_(e1, e2);

    return {
      caseKey: String(row[cKey] || "").trim(),
      injuryName,
      injuryDate,
      firstDate,
      startDate,
      endDate,
    };
  }

  return { case1: build(1), case2: build(2) };
}

function V3TR_minDate_(a, b) {
  const da = (a instanceof Date) ? a : null;
  const db = (b instanceof Date) ? b : null;
  if (!da && !db) return "";
  if (da && !db) return da;
  if (!da && db) return db;
  return (da.getTime() <= db.getTime()) ? da : db;
}
function V3TR_maxDate_(a, b) {
  const da = (a instanceof Date) ? a : null;
  const db = (b instanceof Date) ? b : null;
  if (!da && !db) return "";
  if (da && !db) return da;
  if (!da && db) return db;
  return (da.getTime() >= db.getTime()) ? da : db;
}

/** ===== 施術明細：月次集計（case1/case2 / 確定列参照） ===== */
function V3TR_aggregateDetailMonthly_(shDetail, patientId, start, end) {
  const C = V3TR.CONFIG;
  const map = V3TR_buildHeaderMap_(shDetail);
  const v = shDetail.getDataRange().getValues();
  if (v.length < 2) return { case1: V3TR_emptyAgg_(), case2: V3TR_emptyAgg_() };

  const col = (n) => V3TR_mustCol_(map, n, "施術明細");

  const cPid = col(C.detailCols.patientId);
  const cDt  = col(C.detailCols.treatDate);
  const cCK  = col(C.detailCols.caseKey);

  const cBase = col(C.detailCols.baseFixed);
  const cCold = col(C.detailCols.coldFixed);
  const cWarm = col(C.detailCols.warmFixed);
  const cElec = col(C.detailCols.elecFixed);
  const cRowT = col(C.detailCols.rowTotalFixed);

  const agg1 = V3TR_emptyAgg_();
  const agg2 = V3TR_emptyAgg_();

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;
    const dt = v[r][cDt];
    if (!V3TR_inRange_(dt, start, end)) continue;

    const caseKey = String(v[r][cCK] || "").trim();
    const no = V3TR_caseNoFromCaseKey_(caseKey);
    if (no !== 1 && no !== 2) continue;

    const base = V3TR_num_(v[r][cBase]);
    const cold = V3TR_num_(v[r][cCold]);
    const warm = V3TR_num_(v[r][cWarm]);
    const elec = V3TR_num_(v[r][cElec]);
    const rowT = V3TR_num_(v[r][cRowT]);

    const tgt = (no === 1) ? agg1 : agg2;
    const dk = V3TR_dateKey_(dt);
    if (dk && !tgt._daySet.has(dk)) tgt._daySet.add(dk);

    // 回数は「金額>0の distinct日」
    if (base > 0) { tgt.koryoSum += base; tgt._koryoDaySet.add(dk || ("r" + r)); }
    if (cold > 0) { tgt.coldSum += cold; tgt._coldDaySet.add(dk || ("r" + r)); }
    if (warm > 0) { tgt.warmSum += warm; tgt._warmDaySet.add(dk || ("r" + r)); }
    if (elec > 0) { tgt.elecSum += elec; tgt._elecDaySet.add(dk || ("r" + r)); }

    // case計の確実性を上げたい場合：行合計_確定も合算して保持（任意）
    // → 今回は申請書の内訳（後療/冷/温/電）で case計を作るので未使用だが残しておく
    tgt._rowTotalSum += rowT;
  }

  agg1.visitDays = agg1._daySet.size;
  agg2.visitDays = agg2._daySet.size;

  agg1.koryoCount = agg1._koryoDaySet.size;
  agg1.coldCount  = agg1._coldDaySet.size;
  agg1.warmCount  = agg1._warmDaySet.size;
  agg1.elecCount  = agg1._elecDaySet.size;

  agg2.koryoCount = agg2._koryoDaySet.size;
  agg2.coldCount  = agg2._coldDaySet.size;
  agg2.warmCount  = agg2._warmDaySet.size;
  agg2.elecCount  = agg2._elecDaySet.size;

  return { case1: agg1, case2: agg2 };
}

function V3TR_emptyAgg_() {
  return {
    koryoSum: 0, coldSum: 0, warmSum: 0, elecSum: 0,
    koryoCount: 0, coldCount: 0, warmCount: 0, elecCount: 0,
    visitDays: 0,
    _daySet: new Set(),
    _koryoDaySet: new Set(),
    _coldDaySet: new Set(),
    _warmDaySet: new Set(),
    _elecDaySet: new Set(),
    _rowTotalSum: 0,
  };
}
function V3TR_caseNoFromCaseKey_(caseKey) {
  // 新形式 _C1/_C2 と旧形式 |C1/|C2 の両方を認識
  const m = String(caseKey || "").match(/(?:\|C|_C)([12])$/);
  return m ? Number(m[1]) : 0;
}
function V3TR_num_(v) {
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function V3TR_countKubunInCases_(shCases, patientId, start, end) {
  const C = V3TR.CONFIG;
  const map = V3TR_buildHeaderMap_(shCases);
  const v = shCases.getDataRange().getValues();
  if (v.length < 2) return { initCount: 0, reCount: 0 };

  const cPid = V3TR_mustCol_(map, C.caseCols.patientId, "来院ケース");
  const cDt  = V3TR_mustCol_(map, C.caseCols.treatDate, "来院ケース");
  const cKb  = V3TR_mustCol_(map, C.caseCols.kubun, "来院ケース");

  let initCount = 0;
  let reCount = 0;

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;
    const dt = v[r][cDt];
    if (!V3TR_inRange_(dt, start, end)) continue;
    const k = String(v[r][cKb] || "").trim();
    if (k === "初検") initCount++;
    else if (k === "再検") reCount++;
  }
  return { initCount, reCount };
}

function V3TR_buildCaseMoneyBlock_(agg) {
  const koryoSum = Math.round(agg.koryoSum);
  const coldSum  = Math.round(agg.coldSum);
  const warmSum  = Math.round(agg.warmSum);
  const elecSum  = Math.round(agg.elecSum);

  const koryoCount = agg.koryoCount || 0;
  const koryoUnit = (koryoCount > 0) ? Math.round(koryoSum / koryoCount) : 0;

  const caseTotal = koryoSum + coldSum + warmSum + elecSum;

  return {
    koryoUnit,
    koryoCount,
    koryoSum,

    coldCount: agg.coldCount || 0,
    coldSum,

    warmCount: agg.warmCount || 0,
    warmSum,

    elecCount: agg.elecCount || 0,
    elecSum,

    caseTotal
  };
}

function V3TR_normRatio_(raw) {
  if (raw === "" || raw == null) return 0;
  const n = (typeof raw === "number") ? raw : Number(String(raw).trim());
  if (!isFinite(n)) return 0;
  return (n > 1) ? (n / 100) : n;
}
function V3TR_pickBurdenDigit_(master) {
  const d = master.burdenDigit;
  const dn = Number(d);
  if (isFinite(dn) && dn > 0) return dn;

  const r = V3TR_normRatio_(master.burdenRatio);
  if (!isFinite(r) || r <= 0) return "";
  return Math.round(r * 10);
}

function V3TR_roundToUnit_(value, unit) {
  const v = Number(value);
  const u = Number(unit || 1);
  if (!isFinite(v)) return 0;
  if (!isFinite(u) || u <= 0) return Math.round(v);
  return Math.round(v / u) * u;
}

function V3TR_upsertTransferRows_(shTransfer, rows) {
  const C = V3TR.CONFIG;
  const headers = C.transferCols;
  const map = V3TR_buildHeaderMap_(shTransfer);

  const cKey = V3TR_mustCol_(map, "recordKey", "申請書_転記データ");
  const data = shTransfer.getDataRange().getValues();
  const existing = new Map();
  for (let r = 1; r < data.length; r++) {
    const k = String(data[r][cKey] || "").trim();
    if (k) existing.set(k, r + 1);
  }

  let upserted = 0;

  for (const obj of rows) {
    const key = String(obj.recordKey || "").trim();
    if (!key) continue;

    const rowArr = new Array(headers.length).fill("");
    headers.forEach((h, i) => {
      rowArr[i] = (obj[h] !== undefined) ? obj[h] : "";
    });

    const hitRow = existing.get(key);
    if (hitRow) {
      shTransfer.getRange(hitRow, 1, 1, headers.length).setValues([rowArr]);
    } else {
      shTransfer.getRange(shTransfer.getLastRow() + 1, 1, 1, headers.length).setValues([rowArr]);
      existing.set(key, shTransfer.getLastRow());
    }
    upserted++;
  }

  return upserted;
}