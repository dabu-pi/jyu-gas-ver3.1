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
    initSupport: "初検時相談支援料",
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
    partOrder: "部位順位",
    bui: "部位",
    byomei: "傷病",
    injuryDateFixed: "受傷日_確定",
    coefFixed: "係数",

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

    // 負傷名(1)-(5)  ※申請書行26-30
    "負傷名1", "負傷年月日1", "初検年月日1", "施術開始年月日1", "施術終了年月日1", "実日数1",
    "負傷名2", "負傷年月日2", "初検年月日2", "施術開始年月日2", "施術終了年月日2", "実日数2",

    // 初検料・再検料  ※申請書行33-34
    "初検料_月額", "初検時相談支援料_月額", "再検料_月額", "基本3項目_計",

    // 施療料(1)-(5)  ※申請書行35
    "施療料1", "施療料2", "施療料_計",

    // 部位別明細⑴-⑵  ※申請書行38-43
    "部位1_逓減率", "部位1_後療料_単価", "部位1_後療料_回数", "部位1_後療料_金額",
    "部位1_冷罨法_回数", "部位1_冷罨法_金額",
    "部位1_温罨法_回数", "部位1_温罨法_金額",
    "部位1_電療_回数", "部位1_電療_金額",
    "部位1_計",

    "部位2_逓減率", "部位2_後療料_単価", "部位2_後療料_回数", "部位2_後療料_金額",
    "部位2_冷罨法_回数", "部位2_冷罨法_金額",
    "部位2_温罨法_回数", "部位2_温罨法_金額",
    "部位2_電療_回数", "部位2_電療_金額",
    "部位2_計",

    // 後療料・冷温電 ケース合計（旧互換）
    "後療料_単価", "後療料_回数", "後療料_計",
    "冷罨法_回数", "冷罨法_金額",
    "温罨法_回数", "温罨法_金額",
    "電療_回数", "電療_金額",
    "case計",

    "当月合計", "窓口負担額", "請求金額"
  ],

  /**
   * ★ 療養費支給申請書（新 様式第5号）セルマッピング
   *
   * テンプレシート: 66行×159列(DV)、490結合セル
   * セル番地はテンプレの結合セル左上を指定する。
   *
   * ★注意事項:
   * - 保険者番号: 4列結合×8セルに1桁ずつ
   * - 負傷年月日等の年: 和暦年（令和8→8）
   * - 初検料/再検料/施療料: ラベル内テキストに金額埋め込み
   * - 合計欄(行44-46): 4列結合×6セルに1桁ずつ右詰め
   *
   * 部位行: ⑴行38, ⑵行39, ⑶行40, ⑷行42
   *   1ケース最大2部位→ case1→行38/39, case2→行40/42
   */
  appCellMap: {
    templateSheet: "新　様式第5号",

    // --- 保険者番号: 1桁ずつ（4列結合×8セル） ---
    保険者番号: ["CQ4","CU4","CY4","DC4","DG4","DK4","DO4","DS4"],
    記号:       "BZ5",      // BZ5:CJ7
    番号:       "CK5",      // CK5:DV7

    // --- 被保険者 ---
    被保険者氏名:  "X14",
    住所:          "BF14",

    // --- 受療者 ---
    患者氏名:   "E21",
    生年月日_元号: "AP21",  // 1明 2大 3昭 4平 5令
    生年月日_年: "AY23",

    // --- 負傷の原因 ---
    負傷原因: "BR20",

    // --- 負傷名(1)-(5) 行26-30 ---
    //  ★年は和暦年で書き込む
    //  各行: 負傷名=E{r}:AM{r}, 負傷年月日=AN{r}/AS{r}/AY{r}(和暦年・月・日)
    //        初検年月日=BD{r}/BI{r}/BO{r}, 施術開始=BT{r}/BY{r}/CE{r}
    //        施術終了=CJ{r}/CO{r}/CU{r}, 実日数=CZ{r}:DG{r}
    負傷名: [
      { row: 26, name: "E26", injY: "AN26", injM: "AS26", injD: "AY26",
        iniY: "BD26", iniM: "BI26", iniD: "BO26",
        stY: "BT26", stM: "BY26", stD: "CE26",
        edY: "CJ26", edM: "CO26", edD: "CU26",
        days: "CZ26" },
      { row: 27, name: "E27", injY: "AN27", injM: "AS27", injD: "AY27",
        iniY: "BD27", iniM: "BI27", iniD: "BO27",
        stY: "BT27", stM: "BY27", stD: "CE27",
        edY: "CJ27", edM: "CO27", edD: "CU27",
        days: "CZ27" },
      { row: 28, name: "E28", injY: "AN28", injM: "AS28", injD: "AY28",
        iniY: "BD28", iniM: "BI28", iniD: "BO28",
        stY: "BT28", stM: "BY28", stD: "CE28",
        edY: "CJ28", edM: "CO28", edD: "CU28",
        days: "CZ28" },
      { row: 29, name: "E29", injY: "AN29", injM: "AS29", injD: "AY29",
        iniY: "BD29", iniM: "BI29", iniD: "BO29",
        stY: "BT29", stM: "BY29", stD: "CE29",
        edY: "CJ29", edM: "CO29", edD: "CU29",
        days: "CZ29" },
      { row: 30, name: "E30", injY: "AN30", injM: "AS30", injD: "AY30",
        iniY: "BD30", iniM: "BI30", iniD: "BO30",
        stY: "BT30", stM: "BY30", stD: "CE30",
        edY: "CJ30", edM: "CO30", edD: "CU30",
        days: "CZ30" },
    ],

    // --- 施術日カレンダー（行32: 1-31日） ---
    施術日開始列: "M32",    // M32が1日目、以降4列ごと

    // --- 金額欄 行33-34 ---
    // ★ラベル内テキストに金額埋め込み（例: "初検料1,460円"）
    初検料:         { cell: "E33",  tmpl: "初検料{amt}円" },
    初検時相談支援料: { cell: "Y33", tmpl: "初検時相談\n支援料{amt}円" },
    再検料:         { cell: "Y34",  tmpl: "再検料{amt}円" },
    基本3項目_計:   { cell: "DC33", tmpl: "{amt}円" },   // DC33:DV34

    // --- 施療料 行35 ---
    // ★ラベル内テキストに金額埋め込み（例: "(1)  820円"）
    施療料: [
      { cell: "AC35", no: 1 },   // (1) AC35:AQ35
      { cell: "AR35", no: 2 },   // (2) AR35:BF35
      { cell: "BG35", no: 3 },   // (3) BG35:BU35
      { cell: "BV35", no: 4 },   // (4) BV35:CJ35
      { cell: "CK35", no: 5 },   // (5) CK35:CY35
    ],
    施療料_計: { cell: "DC35", tmpl: "{amt}円" },

    // --- 部位別明細 行38-43 ---
    //  各行の列構造（結合セル左上）:
    //    部位番号=E, 逓減%=H, 逓減開始日=M,
    //    後療料単価=V, 後療料回数=AC, 後療料金額=AH,
    //    冷罨法回数=AR, 冷罨法金額=AW,
    //    温罨法回数=BF, 温罨法金額=BK,
    //    電療回数=BT, 電療金額=BY,
    //    計=CH, 多部位=CR, 多部位後計=CW, 長期=DF,
    //    部位総額=DK（後療+冷+温+電 の合計）
    部位行: [
      { row: 38, label: "E38", teiRate: "H38", teiStart: "M38",
        koryoUnit: "V38", koryoCnt: "AC38", koryoAmt: "AH38",
        coldCnt: "AR38", coldAmt: "AW38",
        warmCnt: "BF38", warmAmt: "BK38",
        elecCnt: "BT38", elecAmt: "BY38",
        subtotal: "CH38", multiCoef: "CR38", multiTotal: "CW38",
        longCoef: "DF38", longTotal: "DK38" },
      { row: 39, label: "E39", teiRate: "H39", teiStart: "M39",
        koryoUnit: "V39", koryoCnt: "AC39", koryoAmt: "AH39",
        coldCnt: "AR39", coldAmt: "AW39",
        warmCnt: "BF39", warmAmt: "BK39",
        elecCnt: "BT39", elecAmt: "BY39",
        subtotal: "CH39", multiCoef: "CR39", multiTotal: "CW39",
        longCoef: "DF39", longTotal: "DK39" },
      { row: 40, label: "E40", teiRate: "H40", teiStart: "M40",
        koryoUnit: "V40", koryoCnt: "AC40", koryoAmt: "AH40",
        coldCnt: "AR40", coldAmt: "AW40",
        warmCnt: "BF40", warmAmt: "BK40",
        elecCnt: "BT40", elecAmt: "BY40",
        subtotal: "CH40", multiCoef: "CR40", multiTotal: "CW40",
        longCoef: "DF40", longTotal: "DK40" },
      { row: 42, label: "E42", teiRate: "H42", teiStart: "M42",
        koryoUnit: "V42", koryoCnt: "AC42", koryoAmt: "AH42",
        coldCnt: "AR42", coldAmt: "AW42",
        warmCnt: "BF42", warmAmt: "BK42",
        elecCnt: "BT42", elecAmt: "BY42",
        subtotal: "CH42", multiCoef: "CR42", multiTotal: "CW42",
        longCoef: "DF42", longTotal: "DK42" },
    ],

    // --- 合計欄 行44-46 ---
    // ★1桁ずつ右詰め（4列結合×6セル + DT{r}="円"固定）
    合計:     ["CV44","CZ44","DD44","DH44","DL44","DP44"],
    一部負担金: ["CV45","CZ45","DD45","DH45","DL45","DP45"],
    請求金額:   ["CV46","CZ46","DD46","DH46","DL46","DP46"],

    // --- 請求区分 行31 ---
    請求区分: "DH31",       // "新規 ・ 継続"

    // --- 摘要 行44-46左 ---
    摘要: "E44",            // E44:CG46
  },
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

    // 部位別集計データ
    const agg = (caseNo === 1) ? detailAgg.case1 : detailAgg.case2;
    const p1 = agg.parts[1] || V3TR_emptyPartAgg_();
    const p2 = agg.parts[2] || V3TR_emptyPartAgg_();

    // 負傷名(1),(2)  ※部位別
    row["負傷名1"] = V3TR_buildInjuryLabel_(p1);
    row["負傷年月日1"] = p1.injuryDate || "";
    row["初検年月日1"] = cs.firstDate || "";
    row["施術開始年月日1"] = cs.startDate || "";
    row["施術終了年月日1"] = cs.endDate || "";
    row["実日数1"] = (caseNo === 1) ? detailAgg.case1.visitDays : detailAgg.case2.visitDays;

    row["負傷名2"] = V3TR_buildInjuryLabel_(p2);
    row["負傷年月日2"] = p2.injuryDate || "";
    row["初検年月日2"] = "";
    row["施術開始年月日2"] = "";
    row["施術終了年月日2"] = "";
    row["実日数2"] = "";

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

    // 施療料(1),(2)  ※初検日の基本料
    row["施療料1"] = p1.shoryoFee || "";
    row["施療料2"] = p2.shoryoFee || "";
    row["施療料_計"] = (p1.shoryoFee || 0) + (p2.shoryoFee || 0) || "";

    // 部位別明細⑴
    row["部位1_逓減率"] = p1.coef;
    row["部位1_後療料_単価"] = p1.koryoUnit;
    row["部位1_後療料_回数"] = p1.koryoCount;
    row["部位1_後療料_金額"] = p1.koryoSum;
    row["部位1_冷罨法_回数"] = p1.coldCount;
    row["部位1_冷罨法_金額"] = p1.coldSum;
    row["部位1_温罨法_回数"] = p1.warmCount;
    row["部位1_温罨法_金額"] = p1.warmSum;
    row["部位1_電療_回数"] = p1.elecCount;
    row["部位1_電療_金額"] = p1.elecSum;
    row["部位1_計"] = p1.partTotal;

    // 部位別明細⑵
    row["部位2_逓減率"] = p2.coef;
    row["部位2_後療料_単価"] = p2.koryoUnit;
    row["部位2_後療料_回数"] = p2.koryoCount;
    row["部位2_後療料_金額"] = p2.koryoSum;
    row["部位2_冷罨法_回数"] = p2.coldCount;
    row["部位2_冷罨法_金額"] = p2.coldSum;
    row["部位2_温罨法_回数"] = p2.warmCount;
    row["部位2_温罨法_金額"] = p2.warmSum;
    row["部位2_電療_回数"] = p2.elecCount;
    row["部位2_電療_金額"] = p2.elecSum;
    row["部位2_計"] = p2.partTotal;

    // ケース合算（旧互換）
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
  const cPart = col(C.detailCols.partOrder);
  const cCoef = col(C.detailCols.coefFixed);
  const cKubun = col(C.detailCols.kubun);
  const cBui = map[C.detailCols.bui];        // optional
  const cByomei = map[C.detailCols.byomei];   // optional
  const cInjDate = map[C.detailCols.injuryDateFixed]; // optional

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
    const partOrder = V3TR_num_(v[r][cPart]) || 1;
    const coef = V3TR_num_(v[r][cCoef]) || 1;
    const kubun = String(v[r][cKubun] || "").trim();

    const tgt = (no === 1) ? agg1 : agg2;
    const dk = V3TR_dateKey_(dt);
    if (dk && !tgt._daySet.has(dk)) tgt._daySet.add(dk);

    // ケース合算（旧互換）
    if (base > 0) { tgt.koryoSum += base; tgt._koryoDaySet.add(dk || ("r" + r)); }
    if (cold > 0) { tgt.coldSum += cold; tgt._coldDaySet.add(dk || ("r" + r)); }
    if (warm > 0) { tgt.warmSum += warm; tgt._warmDaySet.add(dk || ("r" + r)); }
    if (elec > 0) { tgt.elecSum += elec; tgt._elecDaySet.add(dk || ("r" + r)); }
    tgt._rowTotalSum += rowT;

    // 部位別集計
    if (!tgt.parts[partOrder]) {
      tgt.parts[partOrder] = V3TR_emptyPartAgg_();
      // 部位情報は初回のみ設定
      tgt.parts[partOrder].coef = coef;
      tgt.parts[partOrder].bui = (cBui !== undefined) ? String(v[r][cBui] || "").trim() : "";
      tgt.parts[partOrder].byomei = (cByomei !== undefined) ? String(v[r][cByomei] || "").trim() : "";
      tgt.parts[partOrder].injuryDate = (cInjDate !== undefined) ? v[r][cInjDate] : "";
    }
    const ptgt = tgt.parts[partOrder];
    if (base > 0) { ptgt.koryoSum += base; ptgt._koryoDaySet.add(dk || ("r" + r)); }
    if (cold > 0) { ptgt.coldSum += cold; ptgt._coldDaySet.add(dk || ("r" + r)); }
    if (warm > 0) { ptgt.warmSum += warm; ptgt._warmDaySet.add(dk || ("r" + r)); }
    if (elec > 0) { ptgt.elecSum += elec; ptgt._elecDaySet.add(dk || ("r" + r)); }

    // 初検時の施療料を記録
    if (kubun === "初検" && base > 0) {
      ptgt.shoryoFee = base;
    }
  }

  // ケース合算の集計確定
  [agg1, agg2].forEach(function(agg) {
    agg.visitDays = agg._daySet.size;
    agg.koryoCount = agg._koryoDaySet.size;
    agg.coldCount  = agg._coldDaySet.size;
    agg.warmCount  = agg._warmDaySet.size;
    agg.elecCount  = agg._elecDaySet.size;

    // 部位別集計の確定
    Object.keys(agg.parts).forEach(function(key) {
      const p = agg.parts[key];
      p.koryoCount = p._koryoDaySet.size;
      p.coldCount  = p._coldDaySet.size;
      p.warmCount  = p._warmDaySet.size;
      p.elecCount  = p._elecDaySet.size;
      p.koryoUnit  = (p.koryoCount > 0) ? Math.round(p.koryoSum / p.koryoCount) : 0;
      p.partTotal  = p.koryoSum + p.coldSum + p.warmSum + p.elecSum;
    });
  });

  return { case1: agg1, case2: agg2 };
}

function V3TR_emptyAgg_() {
  return {
    koryoSum: 0, coldSum: 0, warmSum: 0, elecSum: 0,
    koryoCount: 0, coldCount: 0, warmCount: 0, elecCount: 0,
    visitDays: 0,
    parts: {},   // { [partOrder]: V3TR_emptyPartAgg_() }
    _daySet: new Set(),
    _koryoDaySet: new Set(),
    _coldDaySet: new Set(),
    _warmDaySet: new Set(),
    _elecDaySet: new Set(),
    _rowTotalSum: 0,
  };
}
function V3TR_emptyPartAgg_() {
  return {
    koryoSum: 0, coldSum: 0, warmSum: 0, elecSum: 0,
    koryoCount: 0, coldCount: 0, warmCount: 0, elecCount: 0,
    koryoUnit: 0, partTotal: 0,
    coef: 1, bui: "", byomei: "", injuryDate: "", shoryoFee: 0,
    _koryoDaySet: new Set(),
    _coldDaySet: new Set(),
    _warmDaySet: new Set(),
    _elecDaySet: new Set(),
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

/** 部位集計から負傷名ラベルを生成（例: "右足関節 捻挫"） */
function V3TR_buildInjuryLabel_(partAgg) {
  if (!partAgg) return "";
  const bui = partAgg.bui || "";
  const byomei = partAgg.byomei || "";
  const label = (bui + " " + byomei).trim();
  return label || "";
}

/* =======================================================================
   申請書転記（テンプレシートへの書き込み）
   ======================================================================= */

/**
 * メニュー：申請書テンプレートへ転記
 * 患者ID + 対象月 を入力→転記データ生成→テンプレシートに書き込み
 */
function V3TR_menuWriteApplication() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const r1 = ui.prompt("申請書転記", "患者ID を入力してください（例：P0001）", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const patientId = (r1.getResponseText() || "").trim();
  if (!patientId) return ui.alert("患者IDが空です。");

  const r2 = ui.prompt("対象月", "対象月（yyyy-MM）を入力してください（例：2026-02）", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const ym = (r2.getResponseText() || "").trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) return ui.alert("形式が違います。yyyy-MM で入力してください。");

  // 1. 転記データ生成(upsert)
  const buildResult = V3TR_buildTransferDataForMonth_(ss, patientId, ym);

  // 2. 転記データシートからcaseNo=1行を取得してテンプレに書き込み
  const shTransfer = ss.getSheetByName(V3TR.CONFIG.sheetNames.transfer);
  const recordKey1 = patientId + "|" + ym + "|C1";
  const recordKey2 = patientId + "|" + ym + "|C2";

  const tMap = V3TR_buildHeaderMap_(shTransfer);
  const tData = shTransfer.getDataRange().getValues();
  const cRK = V3TR_mustCol_(tMap, "recordKey", "申請書_転記データ");

  let row1 = null, row2 = null;
  for (let r = 1; r < tData.length; r++) {
    const k = String(tData[r][cRK] || "").trim();
    if (k === recordKey1) row1 = V3TR_rowToObj_(tData[r], tMap);
    if (k === recordKey2) row2 = V3TR_rowToObj_(tData[r], tMap);
  }

  if (!row1) return ui.alert("転記データが見つかりません: " + recordKey1);

  const written = V3TR_writeToApplication_(ss, row1, row2);

  ui.alert(
    "申請書転記完了\n" +
    "患者: " + patientId + " / " + ym + "\n" +
    "書込セル数: " + written
  );
}

/** 転記データの1行を {列名: 値} オブジェクトに変換 */
function V3TR_rowToObj_(rowArr, headerMap) {
  const obj = {};
  for (const name in headerMap) {
    obj[name] = rowArr[headerMap[name]];
  }
  return obj;
}

/**
 * 申請書テンプレートシートに転記データを書き込む
 *
 * @param {Spreadsheet} ss
 * @param {Object} row1 - caseNo=1 の転記データ行（列名→値）
 * @param {Object|null} row2 - caseNo=2 の転記データ行（なければnull）
 * @return {number} 書き込んだセル数
 */
function V3TR_writeToApplication_(ss, row1, row2) {
  const CM = V3TR.CONFIG.appCellMap;
  const sh = ss.getSheetByName(CM.templateSheet);
  if (!sh) throw new Error("テンプレシート「" + CM.templateSheet + "」が見つかりません。");

  let count = 0;

  /** セルに値を書き込むヘルパー（空・0は書かない） */
  function put(cell, val) {
    if (val === "" || val === null || val === undefined) return;
    sh.getRange(cell).setValue(val);
    count++;
  }

  /** 数値のみ書き込むヘルパー（0は書かない） */
  function putNum(cell, val) {
    const n = Number(val);
    if (!isFinite(n) || n === 0) return;
    sh.getRange(cell).setValue(n);
    count++;
  }

  // ===== 保険者情報 =====
  put(CM.保険者番号, row1["保険者番号"]);
  put(CM.記号, row1["記号"]);
  put(CM.番号, row1["番号"]);

  // ===== 被保険者 =====
  put(CM.被保険者氏名, row1["被保険者氏名"]);
  put(CM.住所, row1["住所"]);

  // ===== 受療者 =====
  put(CM.患者氏名, row1["患者氏名"]);

  // 生年月日 → 元号・年・月・日
  const bd = row1["患者生年月日"];
  if (bd) {
    const bdDate = (bd instanceof Date) ? bd : new Date(bd);
    if (!isNaN(bdDate.getTime())) {
      const era = V3TR_toWareki_(bdDate);
      put(CM.生年月日_元号, era.code);
      put(CM.生年月日_年, era.year);
    }
  }

  // ===== 負傷名(1)-(5): case1部位1, case1部位2, case2部位1, case2部位2 =====
  const injRows = CM.負傷名;
  const injData = V3TR_buildInjuryRows_(row1, row2);

  for (let i = 0; i < injData.length && i < injRows.length; i++) {
    const d = injData[i];
    const m = injRows[i];
    if (!d.name) continue;

    put(m.name, d.name);
    V3TR_putDateYMD_(sh, m.injY, m.injM, m.injD, d.injuryDate);
    V3TR_putDateYMD_(sh, m.iniY, m.iniM, m.iniD, d.firstDate);
    V3TR_putDateYMD_(sh, m.stY, m.stM, m.stD, d.startDate);
    V3TR_putDateYMD_(sh, m.edY, m.edM, m.edD, d.endDate);
    putNum(m.days, d.days);
    count += V3TR_countDateWrites_(d);
  }

  // ===== 初検料・再検料 行33-34 =====
  putNum(CM.初検料, row1["初検料_月額"]);
  putNum(CM.初検時相談支援料, row1["初検時相談支援料_月額"]);
  putNum(CM.再検料, row1["再検料_月額"]);
  putNum(CM.基本3項目_計, row1["基本3項目_計"]);

  // ===== 施療料 行35 =====
  const shoryoData = V3TR_buildShoryoArray_(row1, row2);
  const shoryoCells = CM.施療料;
  for (let i = 0; i < shoryoData.length && i < shoryoCells.length; i++) {
    putNum(shoryoCells[i], shoryoData[i]);
  }
  // 施療料計 = case1 + case2
  const shoryoTotal = Number(row1["施療料_計"] || 0) + Number((row2 || {})["施療料_計"] || 0);
  putNum(CM.施療料_計, shoryoTotal);

  // ===== 部位別明細 行38-43 =====
  // case1部位1 → 行38, case1部位2 → 行39, case2部位1 → 行40, case2部位2 → 行42
  const partRows = CM.部位行;
  const partData = V3TR_buildPartDetailArray_(row1, row2);

  for (let i = 0; i < partData.length && i < partRows.length; i++) {
    const d = partData[i];
    const m = partRows[i];
    if (!d.hasData) continue;

    put(m.label, d.label);
    putNum(m.teiRate, d.teiRate);
    putNum(m.koryoUnit, d.koryoUnit);
    putNum(m.koryoCnt, d.koryoCnt);
    putNum(m.koryoAmt, d.koryoAmt);
    putNum(m.coldCnt, d.coldCnt);
    putNum(m.coldAmt, d.coldAmt);
    putNum(m.warmCnt, d.warmCnt);
    putNum(m.warmAmt, d.warmAmt);
    putNum(m.elecCnt, d.elecCnt);
    putNum(m.elecAmt, d.elecAmt);
    putNum(m.subtotal, d.subtotal);
  }

  // ===== 合計欄 行44-46 =====
  putNum(CM.合計, row1["当月合計"]);
  putNum(CM.一部負担金, row1["窓口負担額"]);
  putNum(CM.請求金額, row1["請求金額"]);

  return count;
}

/* ------- 転記用データ組み立てヘルパー ------- */

/**
 * 負傷名行データを組み立てる（最大5行: case1P1, case1P2, case2P1, case2P2）
 */
function V3TR_buildInjuryRows_(row1, row2) {
  const result = [];

  // case1 部位1
  result.push({
    name:       row1["負傷名1"] || "",
    injuryDate: row1["負傷年月日1"] || "",
    firstDate:  row1["初検年月日1"] || "",
    startDate:  row1["施術開始年月日1"] || "",
    endDate:    row1["施術終了年月日1"] || "",
    days:       row1["実日数1"] || "",
  });
  // case1 部位2
  result.push({
    name:       row1["負傷名2"] || "",
    injuryDate: row1["負傷年月日2"] || "",
    firstDate:  row1["初検年月日2"] || "",
    startDate:  row1["施術開始年月日2"] || "",
    endDate:    row1["施術終了年月日2"] || "",
    days:       row1["実日数2"] || "",
  });

  if (row2) {
    // case2 部位1
    result.push({
      name:       row2["負傷名1"] || "",
      injuryDate: row2["負傷年月日1"] || "",
      firstDate:  row2["初検年月日1"] || "",
      startDate:  row2["施術開始年月日1"] || "",
      endDate:    row2["施術終了年月日1"] || "",
      days:       row2["実日数1"] || "",
    });
    // case2 部位2
    result.push({
      name:       row2["負傷名2"] || "",
      injuryDate: row2["負傷年月日2"] || "",
      firstDate:  row2["初検年月日2"] || "",
      startDate:  row2["施術開始年月日2"] || "",
      endDate:    row2["施術終了年月日2"] || "",
      days:       row2["実日数2"] || "",
    });
  }

  return result;
}

/**
 * 施療料配列を組み立て（最大5: case1P1, case1P2, case2P1, case2P2）
 */
function V3TR_buildShoryoArray_(row1, row2) {
  const arr = [];
  arr.push(Number(row1["施療料1"] || 0));
  arr.push(Number(row1["施療料2"] || 0));
  if (row2) {
    arr.push(Number(row2["施療料1"] || 0));
    arr.push(Number(row2["施療料2"] || 0));
  }
  return arr;
}

/**
 * 部位別明細行データを組み立て
 * 行38=case1P1, 行39=case1P2, 行40=case2P1, 行42=case2P2
 */
function V3TR_buildPartDetailArray_(row1, row2) {
  const result = [];

  function buildOne(row, partNo) {
    const pfx = "部位" + partNo + "_";
    const teiRate = row[pfx + "逓減率"];
    const koryoUnit = Number(row[pfx + "後療料_単価"] || 0);
    const koryoCnt  = Number(row[pfx + "後療料_回数"] || 0);
    const koryoAmt  = Number(row[pfx + "後療料_金額"] || 0);
    const coldCnt   = Number(row[pfx + "冷罨法_回数"] || 0);
    const coldAmt   = Number(row[pfx + "冷罨法_金額"] || 0);
    const warmCnt   = Number(row[pfx + "温罨法_回数"] || 0);
    const warmAmt   = Number(row[pfx + "温罨法_金額"] || 0);
    const elecCnt   = Number(row[pfx + "電療_回数"] || 0);
    const elecAmt   = Number(row[pfx + "電療_金額"] || 0);
    const subtotal  = Number(row[pfx + "計"] || 0);
    const hasData   = subtotal > 0 || koryoAmt > 0;

    return {
      hasData:   hasData,
      label:     "⑴⑵⑶⑷⑸".charAt(result.length) || "",
      teiRate:   teiRate,
      koryoUnit: koryoUnit,
      koryoCnt:  koryoCnt,
      koryoAmt:  koryoAmt,
      coldCnt:   coldCnt,
      coldAmt:   coldAmt,
      warmCnt:   warmCnt,
      warmAmt:   warmAmt,
      elecCnt:   elecCnt,
      elecAmt:   elecAmt,
      subtotal:  subtotal,
    };
  }

  // case1 部位1, 部位2
  result.push(buildOne(row1, 1));
  result.push(buildOne(row1, 2));

  if (row2) {
    result.push(buildOne(row2, 1));
    result.push(buildOne(row2, 2));
  }

  return result;
}

/* ------- 日付ヘルパー ------- */

/**
 * 西暦Date → 和暦 { code, year }
 * code: 1=明治 2=大正 3=昭和 4=平成 5=令和
 */
function V3TR_toWareki_(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return { code: "", year: "" };
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();

  // 令和: 2019-05-01〜
  if (y > 2019 || (y === 2019 && (m > 5 || (m === 5 && day >= 1)))) {
    return { code: 5, year: y - 2018 };
  }
  // 平成: 1989-01-08〜
  if (y > 1989 || (y === 1989 && (m > 1 || (m === 1 && day >= 8)))) {
    return { code: 4, year: y - 1988 };
  }
  // 昭和: 1926-12-25〜
  if (y > 1926 || (y === 1926 && (m > 12 || (m === 12 && day >= 25)))) {
    return { code: 3, year: y - 1925 };
  }
  // 大正: 1912-07-30〜
  if (y > 1912 || (y === 1912 && (m > 7 || (m === 7 && day >= 30)))) {
    return { code: 2, year: y - 1911 };
  }
  // 明治
  return { code: 1, year: y - 1867 };
}

/**
 * Date値を年/月/日の3セルに分解して書き込む
 */
function V3TR_putDateYMD_(sh, cellY, cellM, cellD, dateVal) {
  if (!dateVal) return;
  const d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return;

  sh.getRange(cellY).setValue(d.getFullYear());
  sh.getRange(cellM).setValue(d.getMonth() + 1);
  sh.getRange(cellD).setValue(d.getDate());
}

/** 日付書き込みのカウント用 */
function V3TR_countDateWrites_(d) {
  let c = 0;
  if (d.injuryDate) c += 3;
  if (d.firstDate) c += 3;
  if (d.startDate) c += 3;
  if (d.endDate) c += 3;
  return c;
}

/* =======================================================================
   転記データJSON出力（Python連携用）
   ======================================================================= */

/**
 * メニュー：転記データをJSON出力（ログ表示）
 * GAS実行ログからコピー or ScriptProperties に保存
 */
function V3TR_menuExportJson() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const r1 = ui.prompt("JSON出力", "患者ID を入力してください（例：P0001）", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const patientId = (r1.getResponseText() || "").trim();
  if (!patientId) return ui.alert("患者IDが空です。");

  const r2 = ui.prompt("対象月", "対象月（yyyy-MM）を入力してください（例：2026-02）", ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const ym = (r2.getResponseText() || "").trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) return ui.alert("形式が違います。yyyy-MM で入力してください。");

  const json = V3TR_exportTransferJson_(ss, patientId, ym);

  // JSONシートに出力（コピー用）
  let shJson = ss.getSheetByName("_JSON出力");
  if (!shJson) shJson = ss.insertSheet("_JSON出力");
  shJson.clear();
  shJson.getRange("A1").setValue(json);

  ui.alert("_JSON出力シートのA1にJSONを出力しました。\nコピーしてPythonスクリプトで使用してください。");
}

/**
 * 転記データをJSONテキストとして出力
 *
 * @param {Spreadsheet} ss
 * @param {string} patientId
 * @param {string} ym - yyyy-MM
 * @return {string} JSON文字列
 */
function V3TR_exportTransferJson_(ss, patientId, ym) {
  // まず転記データを生成/更新
  V3TR_buildTransferDataForMonth_(ss, patientId, ym);

  const shTransfer = ss.getSheetByName(V3TR.CONFIG.sheetNames.transfer);
  const tMap = V3TR_buildHeaderMap_(shTransfer);
  const tData = shTransfer.getDataRange().getValues();
  const cRK = V3TR_mustCol_(tMap, "recordKey", "申請書_転記データ");

  const recordKey1 = patientId + "|" + ym + "|C1";
  const recordKey2 = patientId + "|" + ym + "|C2";

  let row1 = null, row2 = null;
  for (let r = 1; r < tData.length; r++) {
    const k = String(tData[r][cRK] || "").trim();
    if (k === recordKey1) row1 = V3TR_rowToObj_(tData[r], tMap);
    if (k === recordKey2) row2 = V3TR_rowToObj_(tData[r], tMap);
  }

  // 通院日リスト（施術明細から取得）
  const visitDays = V3TR_collectVisitDays_(ss, patientId, ym);

  // Date→文字列変換
  const result = {
    case1: V3TR_serializeRow_(row1),
    case2: V3TR_serializeRow_(row2),
    visitDays: visitDays,
  };
  return JSON.stringify(result, null, 2);
}

/** Dateオブジェクトを文字列に変換してJSON化可能にする */
function V3TR_serializeRow_(row) {
  if (!row) return null;
  const out = {};
  for (const key in row) {
    const v = row[key];
    if (v instanceof Date) {
      out[key] = Utilities.formatDate(v, "Asia/Tokyo", "yyyy-MM-dd");
    } else {
      out[key] = v;
    }
  }
  return out;
}

/**
 * 施術明細から当月の通院日リスト（日のみ）を収集
 * @return {number[]} 例: [3, 5, 10, 15, 20]
 */
function V3TR_collectVisitDays_(ss, patientId, ym) {
  const C = V3TR.CONFIG;
  const shDetail = ss.getSheetByName(C.sheetNames.detail);
  if (!shDetail) return [];

  const map = V3TR_buildHeaderMap_(shDetail);
  const v = shDetail.getDataRange().getValues();
  if (v.length < 2) return [];

  const cPid = map[C.detailCols.patientId];
  const cDt = map[C.detailCols.treatDate];
  if (cPid === undefined || cDt === undefined) return [];

  const month = V3TR_parseYM_(ym);
  const daySet = new Set();

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;
    const dt = v[r][cDt];
    if (!V3TR_inRange_(dt, month.start, month.end)) continue;
    daySet.add(dt.getDate());
  }

  return Array.from(daySet).sort(function(a, b) { return a - b; });
}