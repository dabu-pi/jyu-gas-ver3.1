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
    history: "初検情報履歴",
  },

  /**
   * 初検情報履歴シートの列名
   * ★ appendInitHistory_V3_（Ver3_core.js）が書き込む正式ヘッダと一致させること
   */
  historyCols: {
    createdAt:      "作成日時",
    patientId:      "患者ID",
    caseKey:        "caseKey",           // Pass-1 正規マッチキー
    caseNo:         "caseNo",
    initDate:       "施術日(初検日)",    // 月末日以前の最新1件を採用
    injuryDatetime: "負傷の日時",
    injuryPlace:    "負傷の場所",
    injuryStatus:   "負傷時の状況",
    initFindings:   "初検時の所見",
    supportContent: "初検時相談支援の内容",
    injuryFixed:    "受傷日_確定",       // 任意列
  },

  setKeys: {
    initFee: "初検料",
    initSupport: "初検時相談支援料",
    reFee: "再検料",
    roundUnit: "窓口端数単位",
    outputFolderId: "出力フォルダID",
    prefectureNo: "都道府県番号",    // U1 CI2書込用（施術機関所在都道府県番号、2桁）
    torokuKigoNo: "登録記号番号",    // U2(CZ2)・下段分割欄(CR51/DK51/DR51)書込用（例: 契2804440-0-0）
    clinicName: "施術所名",          // D5 施術証明欄: 施術所名 → L59
    clinicAddr: "住所",              // D5 施術証明欄: 所在地 → L58
    clinicPractitioner: "施術者氏名", // D5 施術証明欄: 施術者氏名 → L62
  },

  masterCols: {
    patientId: "患者ID",
    name: "氏名",
    birthday: "生年月日",
    gender: "性別",            // "男" or "女"（申請書 AL21/AL23 丸付け用）
    address1: "住所1",
    address2: "住所2",
    relation: "本人・家族の別",
    insuredName: "被保険者氏名",
    burdenRatio: "負担割合",
    burdenRatioDigit: "一部負担金割合", // DP45用（3/2/1）
    // 保険者情報（患者マスタに同居）
    insurerNo: "保険者番号",
    symbol: "記号",
    number: "番号",
    insurerName: "保険者名",
    insuranceType: "保険種別",
  },

  /** 保険者情報シート（別シート運用の場合。患者マスタに同居ならフォールバック） */
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
    tenki1: "転帰_部位1",
    end2: "施術終了日_部位2",
    tenki2: "転帰_部位2",
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

    "患者氏名", "患者生年月日", "性別", "住所", "続柄",
    "被保険者氏名", "保険者番号", "記号", "番号", "保険者名",
    "保険種別",
    "一部負担金割合",

    // 負傷名(1)-(5)  ※申請書行26-30
    "負傷名1", "負傷年月日1", "初検年月日1", "施術開始年月日1", "施術終了年月日1", "実日数1", "転帰1",
    "負傷名2", "負傷年月日2", "初検年月日2", "施術開始年月日2", "施術終了年月日2", "実日数2", "転帰2",

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

    "当月合計", "窓口負担額", "請求金額",

    // 来院区分サマリー（exportHeaderFromCases_V3 と同一ロジックで導出）
    "Mixed区分", "case1要約", "case2要約", "算定区分", "課金理由要約",

    // 初検情報（初検情報履歴シートから取得：対象月末日時点での最新1件）
    "負傷の日時", "負傷の場所", "負傷の状況", "初検時所見", "初検時相談支援内容",
    "初検取得モード",  // "caseKey" | "patientFallback" | "none"

    // 申請書上段・31行目 書込欄（U7〜）
    "請求区分",  // "新規" | "継続" | "" （同月内治癒再発の両方○は将来対応）
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

    // --- 申請書上段 施術機関情報 ---
    都道府県番号: "CI2",    // CI2:CL3（U1: 施術機関所在都道府県番号、2桁）
    施術機関コード: "CZ2",  // CZ2:DV3（U2: 登録記号番号の数字部分 ★暫定運用）
    // --- 単併区分 行8-13 (U4) ---
    単独: "CT8",            // CT8:CY9（固定「単独」→ テキスト"①.単独"に置換）

    // --- 本家区分 行8-13 (U5) ---
    本人: "DB8",    // DB8:DG9  テンプレート"2.本人"（70歳未満・本人）
    六歳: "DB10",   // DB10:DG11 テンプレート"4.六歳"（6歳未満・就学前）
    家族: "DB12",   // DB12:DG13 テンプレート"6.家族"（70歳未満・被扶養者）
    高一: "DH8",    // DH8:DM9  テンプレート"8.高一"（前期高齢者70-74歳・2割負担）
    高7:  "DH12",   // DH12:DM13 テンプレート"0.高7"（前期高齢者70-74歳・3割負担）

    // --- 給付割合 行8-13 (U6) ---
    // テンプレート実値: DP8='10・９' / DP11='８・７'（全角文字）
    // 片側丸付け: 対象数字1文字のみ置換（U5と同方式）
    給付9割:  "DP8",   // DP8:DV10  割合=1: '９'→'⑨' → '10・⑨'
    給付8_7割: "DP11",  // DP11:DV13 割合=2: '８'→'⑧' / 割合=3: '７'→'⑦'

    // --- 下段 登録記号番号 分割欄 行51-52 ---
    // CR49:DV50 はラベル行「登録記号番号」→ 書き込み禁止
    // 入力欄: 左=CR51:DH52 / 中=DK51:DO52 / 右=DR51:DV52
    登録記号番号_左: "CR51",  // 左欄: 例「契2804440」（ハイフン前の部分・先頭文字含む）
    登録記号番号_中: "DK51",  // 中欄: 例「0」（1つ目ハイフン後）
    登録記号番号_右: "DR51",  // 右欄: 例「0」（2つ目ハイフン後）

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
  const shHistory  = ss.getSheetByName(C.sheetNames.history);  // 任意（なければ空文字）

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
  // ★部位別終了日マップ（集計時に終了後の明細行をスキップする二重防御）
  const cs1 = caseSummary.case1, cs2 = caseSummary.case2;
  const endDates = {
    1: { 1: cs1.endDate1 || null, 2: cs1.endDate2 || null },
    2: { 1: cs2.endDate1 || null, 2: cs2.endDate2 || null },
  };
  const detailAgg = V3TR_aggregateDetailMonthly_(shDetail, patientId, start, end, endDates);
  const kubunCount = V3TR_countKubunInCases_(shCases, patientId, start, end);

  const initFee   = settings.initFee * kubunCount.initCount;
  const support   = (kubunCount.initCount > 0) ? settings.initSupport : 0;
  const reFee     = settings.reFee * kubunCount.reCount;
  const base3sum  = initFee + support + reFee;

  const c1 = V3TR_buildCaseMoneyBlock_(detailAgg.case1);
  const c2 = V3TR_buildCaseMoneyBlock_(detailAgg.case2);

  const total = base3sum + c1.caseTotal + c2.caseTotal;

  const burdenRatio = V3TR_normRatio_(master.burdenRatio);
  const copay = V3TR_roundToUnit_(total * burdenRatio, settings.roundUnit);
  const claim = Math.floor(total - copay); // 小数点以下切り捨て

  // 来院区分サマリー（case1行・case2行の両方に同一値を書き込む）
  const _k1 = caseSummary.case1.kubun || "";
  const _k2 = caseSummary.case2.caseKey ? (caseSummary.case2.kubun || "") : "";
  const _isMixed   = !!_k2;
  const _mixedFlag = _isMixed ? "Mixed" : "通常";
  const _case1Summary = _k1 ? "case1:" + _k1 : "case1:なし";

  // case2:初検(抑制) 判定
  //   [A] 施術継続中 Mixed: case1.endDate が空 or >= case2.startDate → 抑制
  //   [B] 治癒後別負傷:     case1.endDate < case2.startDate（厳密）→ 抑制しない
  const _cs1End   = caseSummary.case1.endDate;
  const _cs2Start = caseSummary.case2.startDate;
  const _isPostRecovery = (_cs1End instanceof Date) && (_cs2Start instanceof Date) && (_cs1End < _cs2Start);
  const _initSuppressed = (_k2 === "初検") && _isMixed && !_isPostRecovery;
  const _case2Summary = !_k2                              ? "case2:なし"
    : (_k2 === "初検" && _initSuppressed)                 ? "case2:初検(抑制)"
    : "case2:" + _k2;

  // 算定区分（display列: header 側 chargeKubun と同一ルール）
  // _effInitFee: 抑制フラグを反映した実効初検料（算定区分の判断にのみ使用、金額計算は変えない）
  const _hasKoryo    = (_k1 === "後療" || _k2 === "後療");
  const _effInitFee  = _initSuppressed ? 0 : initFee;
  const _chargeKubun = _effInitFee > 0 ? "初検"
    : reFee > 0                        ? "再検"
    : _hasKoryo                        ? "後療"
    : "算定なし";

  // 課金理由要約（header 側 chargeReason と同一ルール）
  let _chargeReason;
  if (_effInitFee > 0 && !_isMixed) {
    _chargeReason = "初検のみ";
  } else if (_effInitFee > 0 && _isMixed) {
    _chargeReason = "算定可能な初検ありのため初検採用";
  } else if (reFee > 0 && _isMixed && _initSuppressed) {
    _chargeReason = "初検抑制のため再検採用";
  } else if (reFee > 0 && _isMixed && !_initSuppressed) {
    _chargeReason = "再検ありのため再検採用";
  } else if (reFee > 0) {
    _chargeReason = "再検のみ";
  } else if (_hasKoryo && _isMixed && _initSuppressed) {
    _chargeReason = "初検抑制かつ再検対象なし";
  } else if (_hasKoryo) {
    _chargeReason = "後療のみ";
  } else {
    _chargeReason = "算定なし";
  }

  const rowsOut = [];
  for (const caseNo of [1, 2]) {
    const key = `${patientId}|${ym}|C${caseNo}`;

    const cs = (caseNo === 1) ? caseSummary.case1 : caseSummary.case2;
    const cm = (caseNo === 1) ? c1 : c2;

    const jitsunisu = (caseNo === 1) ? detailAgg.case1.visitDays : detailAgg.case2.visitDays;

    // 初検情報（シートがなければ全列空文字）
    const initInfo = shHistory
      ? V3TR_loadInitInfo_(shHistory, patientId, cs.caseKey || "", end)
      : null;

    const row = {};
    row["recordKey"] = key;
    row["患者ID"] = patientId;
    row["対象月"] = ym;
    row["caseNo"] = caseNo;
    row["caseKey"] = cs.caseKey || "";

    row["患者氏名"] = master.name || "";
    row["患者生年月日"] = master.birthday || "";
    row["性別"] = master.gender || "";   // "男" or "女"
    row["住所"] = master.address || "";
    row["続柄"] = master.relation || "";

    row["被保険者氏名"] = master.insuredName || "";
    // 保険者情報: 保険者情報シート優先、なければ患者マスタからフォールバック
    row["保険者番号"] = insurer.insurerNo || master.insurerNo || "";
    row["記号"] = insurer.symbol || master.symbol || "";
    row["番号"] = insurer.number || master.number || "";
    row["保険者名"] = insurer.insurerName || master.insurerName || "";
    row["保険種別"] = master.insuranceType || "";
    row["一部負担金割合"] = V3TR_pickBurdenDigit_(master) || "";

    // 部位別集計データ
    const agg = (caseNo === 1) ? detailAgg.case1 : detailAgg.case2;
    const p1 = agg.parts[1] || V3TR_emptyPartAgg_();
    const p2 = agg.parts[2] || V3TR_emptyPartAgg_();

    // 負傷名(1),(2)  ※部位別
    // 来院ケースに日付がない場合、施術明細の日付範囲からフォールバック
    const aggDates = V3TR_aggDateRange_(agg);
    const p1Dates  = V3TR_aggDateRange_(p1);   // 部位1専用（施術終了年月日の精度向上）
    const p2Dates  = V3TR_aggDateRange_(p2);   // 部位2専用

    // 部位1
    row["負傷名1"] = V3TR_buildInjuryLabel_(p1);
    row["負傷年月日1"] = p1.injuryDate || "";
    row["初検年月日1"] = cs.startDate1 || cs.firstDate || aggDates.minDate || "";
    row["施術開始年月日1"] = cs.startDate1 || cs.startDate || aggDates.minDate || "";
    row["施術終了年月日1"] = cs.endDate1 || p1Dates.maxDate || aggDates.maxDate || "";  // 終了日優先→部位別最終施術日→ケース全体
    row["実日数1"] = p1.visitDays || jitsunisu;  // ★部位別実日数（フォールバック:ケース全体）
    row["転帰1"] = cs.tenki1 || "";

    // 部位2
    row["負傷名2"] = V3TR_buildInjuryLabel_(p2);
    row["負傷年月日2"] = p2.injuryDate || "";
    row["初検年月日2"] = cs.startDate2 || cs.firstDate || aggDates.minDate || "";
    row["施術開始年月日2"] = cs.startDate2 || cs.startDate || aggDates.minDate || "";
    row["施術終了年月日2"] = cs.endDate2 || p2Dates.maxDate || aggDates.maxDate || "";  // 終了日優先→部位別最終施術日→ケース全体
    row["実日数2"] = (V3TR_buildInjuryLabel_(p2)) ? (p2.visitDays || jitsunisu) : "";  // ★部位別実日数
    row["転帰2"] = cs.tenki2 || "";

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

    // 来院区分サマリー（case1行・case2行の両方に同一値）
    row["Mixed区分"]    = _mixedFlag;
    row["case1要約"]    = _case1Summary;
    row["case2要約"]    = _case2Summary;
    row["算定区分"]     = _chargeKubun;
    row["課金理由要約"] = _chargeReason;

    // 初検情報（caseKey 単位で取得。シート未存在 or 行なしは空文字）
    row["負傷の日時"]         = initInfo ? initInfo.injuryDatetime  : "";
    row["負傷の場所"]         = initInfo ? initInfo.injuryPlace      : "";
    row["負傷の状況"]         = initInfo ? initInfo.injuryStatus     : "";
    row["初検時所見"]         = initInfo ? initInfo.initFindings     : "";
    row["初検時相談支援内容"] = initInfo ? initInfo.supportContent   : "";
    row["初検取得モード"]     = initInfo ? initInfo.matchMode        : "none";

    // U7 請求区分: 初検年月日（cs.firstDate）の年月と対象月（ym: "yyyy-MM"）を比較
    // 新規 = 初検月が対象月と同じ / 継続 = 初検月が対象月より前
    // 同月内治癒再発（「新規・継続」両方○）は将来対応。現時点では "新規" or "継続" の単一値。
    {
      const [ymYear, ymMonth] = ym.split("-").map(Number);
      const initD = (cs.firstDate instanceof Date) ? cs.firstDate : null;
      if (initD) {
        const iYear  = initD.getFullYear();
        const iMonth = initD.getMonth() + 1;  // 0-based → 1-based
        row["請求区分"] = (iYear === ymYear && iMonth === ymMonth) ? "新規" : "継続";
      } else {
        row["請求区分"] = "";  // 初検日が不明な場合は空（手入力を促す）
      }
    }

    // D2 継続月数・頻回: 内部値計算のみ。M31への出力は当面停止（B案: 既定で書かない）
    // ★設計確定（2026-03-20）: 正本=摘要欄（手動）+長期欄（頻回→0.5/長期のみ→0.75、手動）
    // ★M31（経過欄）は空欄許容。row["経過"] は常に "" とし Python 側も書かない。
    // ★将来 M31 自動出力を復活させる場合は、下記コメントアウト部を有効化すること。
    //
    // --- 復活時はここから有効化 ---
    // if (caseNo === 1) {
    //   const injD1 = (p1.injuryDate instanceof Date) ? p1.injuryDate
    //               : (cs.startDate1 instanceof Date)  ? cs.startDate1
    //               : (cs.firstDate  instanceof Date)  ? cs.firstDate : null;
    //   if (injD1 && cs.caseKey) {
    //     const d2res = V3TR_calcD2Keizoku_(shCases, patientId, cs.caseKey, injD1, ym);
    //     const kd = d2res.displayMonths;
    //     row["経過"] = (kd !== "" && jitsunisu)
    //       ? kd + "ヶ月 月" + jitsunisu + "回"
    //       : (kd !== "") ? kd + "ヶ月" : "";
    //   } else { row["経過"] = ""; }
    // } else { row["経過"] = ""; }
    // --- ここまで ---
    row["経過"] = ""; // 常に空。M31出力停止中（設計確定 2026-03-20）

    // RC-1修正: case2 データが来院ケース・施術明細の両方に存在しない月は
    // 空レコード（caseKey=""・全金額0）の出力を抑制する。
    // cs.caseKey が非空 or visitDays>0 のいずれかがあれば不整合検出のため出力を維持する。
    if (caseNo === 2 && !cs.caseKey && detailAgg.case2.visitDays === 0) continue;

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

/**
 * 設定シートから施術機関固定情報を読み込む（U1/U2/下段登録記号番号分割欄/D5施術証明欄）
 * 返り値: { prefectureNo, torokuKigoNo, clinicName, clinicAddr, clinicPractitioner }
 */
function V3TR_loadClinicInfo_(shSettings) {
  const C = V3TR.CONFIG;
  if (!shSettings) return { prefectureNo: "", torokuKigoNo: "", clinicName: "", clinicAddr: "", clinicPractitioner: "" };
  const v = shSettings.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < v.length; r++) {
    const k = String(v[r][0] || "").trim();
    if (!k) continue;
    map[k] = String(v[r][1] || "").trim();
  }
  return {
    prefectureNo:       map[C.setKeys.prefectureNo]       || "",
    torokuKigoNo:       map[C.setKeys.torokuKigoNo]       || "",
    clinicName:         map[C.setKeys.clinicName]         || "",
    clinicAddr:         map[C.setKeys.clinicAddr]         || "",
    clinicPractitioner: map[C.setKeys.clinicPractitioner] || "",
  };
}

/**
 * 対象月末日時点の年齢を計算する
 * @param {Date|string} birthday - 生年月日（Date or "YYYY-MM-DD"）
 * @param {string} ym - 対象月 "YYYY-MM"
 * @return {number|null} 年齢。計算できない場合 null
 */
function V3TR_calcAgeAtEndOfMonth_(birthday, ym) {
  if (!birthday || !ym) return null;
  let bd;
  if (birthday instanceof Date) {
    bd = birthday;
  } else {
    const parts = String(birthday).split(/[-\/]/);
    if (parts.length < 3) return null;
    bd = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }
  if (isNaN(bd.getTime())) return null;
  const ymParts = String(ym).split("-");
  if (ymParts.length < 2) return null;
  const y = Number(ymParts[0]);
  const m = Number(ymParts[1]);
  // 対象月末日（month=m は 0-indexed で翌月 0 日 = 当月末日）
  const endOfMonth = new Date(y, m, 0);
  let age = endOfMonth.getFullYear() - bd.getFullYear();
  const monthDiff = endOfMonth.getMonth() - bd.getMonth();
  if (monthDiff < 0 || (monthDiff === 0 && endOfMonth.getDate() < bd.getDate())) {
    age--;
  }
  return (isFinite(age) && age >= 0) ? age : null;
}

/**
 * U5本家区分: row1のデータから appCellMap のキー名を返す（暫定ルール）
 *
 * 判定ロジック（優先順位順）:
 *   1. 保険種別=6（後期高齢）→ "高一"基本。7割給付（負担3割）のみ "高7"（★制度確定2026-03-20）
 *   2. 6歳未満（就学前）→ "六歳"（DB10）
 *   3. 70〜74歳 + 2割負担 → "高一"（DH8）
 *   4. 70〜74歳 + 3割負担 → "高7"（DH12）
 *   5. 70〜74歳 + 割合不明 → "高一"（安全側）
 *   6. 75歳以上 → 後期高齢者扱い。高一基本、7割給付のみ高7
 *   7. 70歳未満 + 続柄="本人" → "本人"（DB8）
 *   8. 70歳未満 + 続柄その他 → "家族"（DB12）
 *
 * 生年月日が不明な場合は年齢区分をスキップして続柄のみで本人/家族判定。
 * 保険種別はGASマスタで"後期高齢"等の名称文字列で保存されているため、名称→数値変換を内部で行う。
 *
 * @param {Object} row1 - transferData の row（患者毎データ）
 * @return {string|null} appCellMap のキー名 or null（書込なし）
 */
function V3TR_deriveHonkeku_(row1) {
  // 保険種別: 数値(6)も名称文字列("後期高齢")も数値に正規化
  // GASマスタは"協会けんぽ"等の文字列で保存されているため、文字列→数値マップが必要
  const INS_TYPE_NAME_MAP_ = {
    "協会けんぽ": 1, "組合": 2, "共済": 3, "国保": 4, "退職": 5, "後期高齢": 6,
  };
  const insuranceTypeRaw = row1["保険種別"];
  const insuranceType = (Number(insuranceTypeRaw) || INS_TYPE_NAME_MAP_[String(insuranceTypeRaw || "").trim()] || 0);
  const relation      = String(row1["続柄"] || "").trim();
  const burden        = Number(row1["一部負担金割合"]) || 0;
  const birthday      = row1["患者生年月日"];
  const ym            = String(row1["対象月"] || "");

  // 後期高齢者（保険種別=6）→ 高一 基本。7割給付（負担3割）のみ 高7
  // ★制度確定（2026-03-20）: 本人/家族区分は使わない。給付割合は U6 側で表現する。
  if (insuranceType === 6) return (burden === 3) ? "高7" : "高一";

  const age = V3TR_calcAgeAtEndOfMonth_(birthday, ym);

  // 6歳未満（就学前）
  if (age !== null && age < 6) return "六歳";

  // 70〜74歳（前期高齢者）
  if (age !== null && age >= 70 && age <= 74) {
    if (burden === 2) return "高一";
    if (burden === 3) return "高7";
    return "高一"; // 負担割合不明は安全側（高一=8割給付）
  }

  // 75歳以上 → 後期高齢者扱い（保険種別=6と同じルール）
  // ★保険種別が6以外でも75歳超は後期高齢者として高一/高7 で判定する
  if (age !== null && age >= 75) return (burden === 3) ? "高7" : "高一";

  // 70歳未満（年齢不明含む）: 続柄で判定
  return (relation === "本人") ? "本人" : "家族";
}

/**
 * 登録記号番号から施術機関コードを導出（暫定ルール）
 * 例: "契2804440-0-0" → "2804440-0-0"
 * ルール: 先頭の「協」または「契」の1文字のみ除去。ハイフンはそのまま保持。
 * ★ 公式一次資料での完全確認未完了。現時点の暫定運用（docs/JREC-01_申請書様式運用メモ.md §4 U2参照）。
 */
function V3TR_deriveClinicCode_(torokuKigoNo) {
  if (!torokuKigoNo) return "";
  let s = String(torokuKigoNo).trim();
  if (s.charAt(0) === "協" || s.charAt(0) === "契") {
    s = s.substring(1);
  }
  return s; // ハイフン保持
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

    // 住所1 + 住所2 を結合
    var addr1 = String(get(C.masterCols.address1) || "").trim();
    var addr2 = String(get(C.masterCols.address2) || "").trim();
    var address = (addr1 + " " + addr2).trim();

    const result = {
      name: get(C.masterCols.name),
      birthday: (get(C.masterCols.birthday) instanceof Date) ? get(C.masterCols.birthday) : "",
      gender: String(get(C.masterCols.gender) || "").trim(),  // "男" or "女"
      address: address,
      relation: get(C.masterCols.relation),
      insuredName: get(C.masterCols.insuredName),
      burdenRatio: get(C.masterCols.burdenRatio),
      burdenDigit: get(C.masterCols.burdenRatioDigit),
      // 保険者情報（患者マスタに同居）
      insurerNo: String(get(C.masterCols.insurerNo) || "").trim(),
      symbol: String(get(C.masterCols.symbol) || "").trim(),
      number: String(get(C.masterCols.number) || "").trim(),
      insurerName: String(get(C.masterCols.insurerName) || "").trim(),
      insuranceType: String(get(C.masterCols.insuranceType) || "").trim(),
    };
    Logger.log("[loadMaster] 取得: patientId=" + patientId + " name=" + result.name + " insuranceType=" + result.insuranceType);
    return result;
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
    const result = {
      insurerNo: String(get(C.insurerCols.insurerNo) || "").trim(),
      symbol: String(get(C.insurerCols.symbol) || "").trim(),
      number: String(get(C.insurerCols.number) || "").trim(),
      insurerName: String(get(C.insurerCols.insurerName) || "").trim(),
    };
    Logger.log("[loadInsurer] 取得: patientId=" + patientId + " insurerNo=" + result.insurerNo + " insurerName=" + result.insurerName);
    return result;
  }
  Logger.log("[loadInsurer] 該当なし: patientId=" + patientId);
  return {};
}

function V3TR_parseYM_(ym) {
  const [y, m] = ym.split("-").map(n => Number(n));
  const start = new Date(y, m - 1, 1);
  const end = new Date(y, m, 1);
  return { start, end };
}

/**
 * D2 継続月数・頻回 計算（申請書 M31 経過欄 補助表示用）
 *
 * caseKey 単位で、起算月（初検日16日以降→翌月起算）からの月別来院日数を集計し、
 * 月10回以上の連続月数（rawContMonths）と頻回開始フラグ（freqStarted）を返す。
 *
 * ★制度上の位置づけ:
 *   - 継続月数の公式記録先: 摘要欄（手動記入）
 *   - 頻回の公式記録先: 申請書「頻回」欄（0.5 を記載。長期欄0.75は書かない）
 *   - M31「経過」欄への記載: 補助表示扱い（義務なし・制度違反でもない）
 * ★ calcMonthsElapsed_V3_ は単純経過月数のため使用しない。
 *
 * @param {Sheet} shCases  - 来院ケースシート（全期間データが必要）
 * @param {string} patientId
 * @param {string} caseKey
 * @param {Date}   injuryDate - 起算月計算用（受傷日または初検日）
 * @param {string} ym         - 対象月 "yyyy-MM"（終端月として使用）
 * @return {{ rawContMonths: number, freqStarted: boolean, displayMonths: number|string }}
 *   displayMonths: 1〜5か月目→実連続月数 / 6か月目以降(freqStarted)→6固定 / 対象外→""
 */
function V3TR_calcD2Keizoku_(shCases, patientId, caseKey, injuryDate, ym) {
  const empty = { rawContMonths: 0, freqStarted: false, displayMonths: "" };
  if (!shCases || !patientId || !caseKey || !(injuryDate instanceof Date)) return empty;

  // 起算月（初検日16日以降→翌月起算）
  let sy = injuryDate.getFullYear();
  let sm = injuryDate.getMonth(); // 0-indexed
  if (injuryDate.getDate() >= 16) {
    sm++;
    if (sm > 11) { sm = 0; sy++; }
  }

  // 対象月（ym の終端月）
  const [ymY, ymM] = ym.split("-").map(Number);
  const ey = ymY, em = ymM - 1; // 0-indexed
  const totalMonths = (ey - sy) * 12 + (em - sm) + 1;
  if (totalMonths <= 0) return empty;

  // 来院ケースシートを読み込み（全期間・患者×caseKey でフィルタ）
  const C  = V3TR.CONFIG;
  const vals = shCases.getDataRange().getValues();
  const hmap = V3TR_buildHeaderMap_(shCases);
  const pidCol = hmap[C.caseCols.patientId];
  const ckCol  = hmap[C.caseCols.caseKey];
  const dtCol  = hmap[C.caseCols.treatDate];
  if (pidCol === undefined || ckCol === undefined || dtCol === undefined) return empty;

  // 月別来院日数（distinct 施術日でカウント）
  const visitSets = {};
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (String(row[pidCol] || "").trim() !== patientId) continue;
    if (String(row[ckCol]  || "").trim() !== caseKey)  continue;
    const td = row[dtCol];
    if (!(td instanceof Date)) continue;
    const idx = (td.getFullYear() - sy) * 12 + (td.getMonth() - sm);
    if (idx < 0 || idx >= totalMonths) continue;
    if (!visitSets[idx]) visitSets[idx] = {};
    // 日付の同一性を yyyymmdd 数値で担保
    const dk = td.getFullYear() * 10000 + (td.getMonth() + 1) * 100 + td.getDate();
    visitSets[idx][dk] = true;
  }
  const counts = {};
  for (const idx in visitSets) {
    counts[idx] = Object.keys(visitSets[idx]).length;
  }

  // 「対象月より前の月」で 5 連続達成（freqStarted）済みかを判定
  let streak = 0;
  let freqStartedBefore = false;
  for (let m = 0; m < totalMonths - 1; m++) {
    const cnt = counts[m] || 0;
    if (cnt >= 10) {
      streak++;
      if (streak >= 5) { freqStartedBefore = true; break; } // 5連続達成 → 以降は「6」固定
    } else {
      streak = 0; // freqStarted 未達のためリセット可
    }
  }

  // 対象月より前に頻回成立 → 対象月は 6か月目以降 → displayMonths = 6 固定
  if (freqStartedBefore) {
    return { rawContMonths: -1, freqStarted: true, displayMonths: 6 };
  }

  // 対象月を処理（当月が10回以上なら連続+1、未満なら0リセット）
  const curCnt = counts[totalMonths - 1] || 0;
  if (curCnt >= 10) {
    streak++;
  } else {
    streak = 0;
  }

  // 対象月が 5か月目完了（freqStarted=true）の場合:
  // 今月は「5」表示（翌月の申請から「6」固定になる）
  return {
    rawContMonths:  streak,
    freqStarted:    (streak >= 5),
    displayMonths:  (streak > 0) ? streak : "",
  };
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
  const cT1 = map[C.caseCols.tenki1] || -1, cT2 = map[C.caseCols.tenki2] || -1;
  const cKubun = col(C.caseCols.kubun);

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
    if (!obj) return {
      caseKey: "", injuryName: "", injuryDate: "",
      firstDate: "", startDate: "", endDate: "",
      startDate1: "", endDate1: "", startDate2: "", endDate2: "",
      tenki1: "", tenki2: "", kubun: "",
    };

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
      // 部位別の日付（部位2用）
      startDate1: s1, endDate1: e1,
      startDate2: s2, endDate2: e2,
      // 転帰
      tenki1: (cT1 >= 0) ? String(row[cT1] || "").trim() : "",
      tenki2: (cT2 >= 0) ? String(row[cT2] || "").trim() : "",
      kubun: String(row[cKubun] || "").trim(),
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

/**
 * 施術明細の集計データ(_daySet)から日付範囲を取得。
 * caseSummary（来院ケース）にデータがない場合のフォールバック用。
 */
function V3TR_aggDateRange_(agg) {
  var minDate = null;
  var maxDate = null;
  if (agg && agg._daySet && agg._daySet.size > 0) {
    agg._daySet.forEach(function(dk) {
      // dk = "yyyy-MM-dd" 形式の文字列（V3TR_dateKey_）
      if (!dk || typeof dk !== "string") return;
      var parts = dk.split("-");
      if (parts.length !== 3) return;
      var dt = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      if (isNaN(dt.getTime())) return;
      if (!minDate || dt.getTime() < minDate.getTime()) minDate = dt;
      if (!maxDate || dt.getTime() > maxDate.getTime()) maxDate = dt;
    });
  }
  return { minDate: minDate || "", maxDate: maxDate || "" };
}

/** ===== 施術明細：月次集計（case1/case2 / 確定列参照） ===== */
function V3TR_aggregateDetailMonthly_(shDetail, patientId, start, end, endDates) {
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

    // ★二重防御: この部位の終了日を超えた明細行はスキップ
    if (endDates) {
      const partEnd = (endDates[no] || {})[partOrder];
      if (partEnd instanceof Date && dt instanceof Date && dt.getTime() > partEnd.getTime()) {
        continue;
      }
    }

    if (dk && !tgt._daySet.has(dk)) tgt._daySet.add(dk);

    // ケース合算（旧互換）★初検日のbaseは施療料なので後療料に含めない
    if (kubun !== "初検" && base > 0) { tgt.koryoSum += base; tgt._koryoDaySet.add(dk || ("r" + r)); }
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
    // ★部位別実日数カウント
    if (dk) ptgt._daySet.add(dk);
    // ★初検日: baseは施療料（shoryoFee）に記録、後療料には含めない
    if (kubun === "初検") {
      if (base > 0) ptgt.shoryoFee = base;
    } else {
      if (base > 0) { ptgt.koryoSum += base; ptgt._koryoDaySet.add(dk || ("r" + r)); }
    }
    if (cold > 0) { ptgt.coldSum += cold; ptgt._coldDaySet.add(dk || ("r" + r)); }
    if (warm > 0) { ptgt.warmSum += warm; ptgt._warmDaySet.add(dk || ("r" + r)); }
    if (elec > 0) { ptgt.elecSum += elec; ptgt._elecDaySet.add(dk || ("r" + r)); }
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
      p.visitDays  = p._daySet.size;  // ★部位別実日数
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

/**
 * 初検情報履歴から対象月末日以前の最新1件を返す。
 *
 * 【マッチ戦略（2パス）】
 *   Pass-1（正規ルート）: caseKey 列が存在 かつ 行の caseKey が引数と完全一致
 *                         → matchMode: "caseKey"
 *   Pass-2（フォールバック）: Pass-1 で行が見つからない場合のみ実行。
 *                             patientId のみでマッチ（caseKey 不問）
 *                             → matchMode: "patientFallback"
 *   行なし: null を返す（呼び出し側が matchMode: "none" として扱う）
 *
 * @param {Sheet}  shHistory  - 初検情報履歴シート
 * @param {string} patientId  - 患者ID
 * @param {string} caseKey    - 対象caseKey（空文字可）
 * @param {Date}   endExcl    - 対象月の翌月1日（exclusive）
 * @returns {{injuryDatetime, injuryPlace, injuryStatus, initFindings, supportContent, matchMode}|null}
 */
function V3TR_loadInitInfo_(shHistory, patientId, caseKey, endExcl) {
  if (!shHistory || shHistory.getLastRow() < 2) return null;
  const C   = V3TR.CONFIG;
  const map = V3TR_buildHeaderMap_(shHistory);
  const v   = shHistory.getDataRange().getValues();

  const cPid  = map[C.historyCols.patientId];
  const cCK   = map[C.historyCols.caseKey];    // 任意列（undefined の場合もある）
  // 旧ヘッダ互換 alias: 新ヘッダ名で見つからない場合に旧名にフォールバック（cDate は null 早期リターンを防ぐため必須）
  const cDate = map[C.historyCols.initDate] !== undefined
                ? map[C.historyCols.initDate] : map["初検日"];
  if (cPid === undefined || cDate === undefined) return null;

  // 日付フィルタ共通: 患者ID一致 かつ 対象月末日以前
  function eligible(r) {
    if (String(v[r][cPid] || "").trim() !== patientId) return false;
    const dt = v[r][cDate];
    if (!(dt instanceof Date)) return false;
    if (endExcl instanceof Date && dt.getTime() >= endExcl.getTime()) return false;
    return true;
  }
  function latest(rows) {
    let best = null;
    for (const r of rows) {
      const dt = v[r][cDate];
      if (!best || dt.getTime() > v[best][cDate].getTime()) best = r;
    }
    return best;
  }

  // Pass-1: caseKey 列が存在し、行の caseKey が対象と完全一致
  let bestRow = null;
  let matchMode = "none";

  if (cCK !== undefined && caseKey) {
    const candidates = [];
    for (let r = 1; r < v.length; r++) {
      if (!eligible(r)) continue;
      if (String(v[r][cCK] || "").trim() === caseKey) candidates.push(r);
    }
    const found = latest(candidates);
    if (found !== null) { bestRow = found; matchMode = "caseKey"; }
  }

  // Pass-2: フォールバック（patientId のみ）
  if (bestRow === null) {
    const candidates = [];
    for (let r = 1; r < v.length; r++) {
      if (eligible(r)) candidates.push(r);
    }
    const found = latest(candidates);
    if (found !== null) { bestRow = found; matchMode = "patientFallback"; }
  }

  if (bestRow === null) return null;

  const row = v[bestRow];
  // ★ Date セルは Utilities.formatDate で変換（String(dateObj) で英語化するのを防ぐ）
  const get = (col) => {
    if (col === undefined) return "";
    const val = row[col];
    if (val instanceof Date) return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
    return String(val || "").trim();
  };
  // 旧ヘッダ名フォールバック（列名が変わった4列のみ alias 解決）
  const a = (newKey, oldKey) => map[newKey] !== undefined ? map[newKey] : map[oldKey];
  return {
    matchMode,
    injuryDatetime: get(map[C.historyCols.injuryDatetime]),
    injuryPlace:    get(map[C.historyCols.injuryPlace]),
    injuryStatus:   get(a(C.historyCols.injuryStatus, "負傷の状況")),
    initFindings:   get(a(C.historyCols.initFindings,  "初検時所見")),
    supportContent: get(a(C.historyCols.supportContent, "初検時相談支援内容")),
  };
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
    koryoUnit: 0, partTotal: 0, visitDays: 0,
    coef: 1, bui: "", byomei: "", injuryDate: "", shoryoFee: 0,
    _daySet: new Set(),
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
  // caseNo / end1 / end2: 存在しない列はインデックス -1 で安全にフォールバック
  const cNo  = map[C.caseCols.caseNo] >= 0 ? map[C.caseCols.caseNo] : -1;
  const cE1  = map[C.caseCols.end1]   >= 0 ? map[C.caseCols.end1]   : -1;
  const cE2  = map[C.caseCols.end2]   >= 0 ? map[C.caseCols.end2]   : -1;

  let rawInitCount = 0;
  let rawReCount   = 0;
  // caseNo → { hasInit, initDate, endDate }
  //   initDate: 当月内 kubun=初検 の最早日
  //   endDate : 当月内 施術終了日_部位1/2 の最遅日（複数行あれば最遅を保持）
  const caseInfo = {};

  for (let r = 1; r < v.length; r++) {
    if (String(v[r][cPid] || "").trim() !== patientId) continue;
    const dt = v[r][cDt];
    if (!V3TR_inRange_(dt, start, end)) continue;
    const k  = String(v[r][cKb] || "").trim();
    const no = cNo >= 0 ? Number(v[r][cNo] || 0) : 1;
    if (!caseInfo[no]) caseInfo[no] = { hasInit: false, initDate: null, endDate: null };

    if (k === "初検") {
      rawInitCount++;
      caseInfo[no].hasInit = true;
      if (!caseInfo[no].initDate || dt < caseInfo[no].initDate) caseInfo[no].initDate = dt;
    } else if (k === "再検") {
      rawReCount++;
    }

    // 施術終了日（部位1/2 で遅い方を保持）
    const e1 = cE1 >= 0 && v[r][cE1] instanceof Date ? v[r][cE1] : null;
    const e2 = cE2 >= 0 && v[r][cE2] instanceof Date ? v[r][cE2] : null;
    const eMax = (e1 && e2) ? (e1 > e2 ? e1 : e2) : (e1 || e2);
    if (eMax && (!caseInfo[no].endDate || eMax > caseInfo[no].endDate)) {
      caseInfo[no].endDate = eMax;
    }
  }

  // ─── 有効な初検料算定件数（validInitCount）を算出 ─────────────────────────────
  // 制度根拠（厚生労働省集団指導資料）:
  //   再検料は「初検料を算定する初検の日後、最初の後療の日のみ」算定可
  //   現に施術継続中に他の負傷が発生した場合、初検料は合わせて1回 → 再検料も増えない [A]
  //   治癒後に同月内で新たな別負傷が発生した場合、初検料を再度算定可 → 再検料も別途1回 [B]
  //
  // 判定: 月内に2つの初検 event が存在する場合、先行ケースの終了日 < 後続ケースの初検日 なら [B]
  //   [A] 施術継続中 Mixed: 終了日なし or 終了日 >= 後続初検日 → validInitCount=1
  //   [B] 治癒後の新規別負傷: 先行ケース終了日 < 後続ケース初検日 → validInitCount=2
  //
  // エッジケース:
  //   earlier.endDate === later.initDate: 厳密 < のため施術継続中 [A] 扱い（同日は保守側に倒す）
  //   earlier.endDate が空（Date でない）: isPostRecovery=false → [A] 扱い
  //   caseNo 列なし: 全行を caseNo=1 扱いにするため initCases.length<=1 → validInitCount=min(rawInitCount,1)
  const initCases = Object.keys(caseInfo).filter(function(no) { return caseInfo[no].hasInit; });
  let validInitCount;
  if (initCases.length <= 1) {
    validInitCount = initCases.length;
  } else {
    // 2件の初検あり → 時系列順で先行/後続を特定し、先行ケースが後続の初検日前に終了しているか判定
    const sorted = initCases.map(function(no) {
      return { initDate: caseInfo[no].initDate, endDate: caseInfo[no].endDate };
    }).sort(function(a, b) { return a.initDate - b.initDate; });
    const earlier = sorted[0];
    const later   = sorted[1];
    // earlier.endDate < later.initDate（厳密）→ 治癒後の新規別負傷 [B]
    // earlier.endDate が Date でない or >= later.initDate → 施術継続中 [A]
    const isPostRecovery = earlier.endDate instanceof Date && earlier.endDate < later.initDate;
    validInitCount = isPostRecovery ? 2 : 1;
  }

  // ★ initCount: validInitCount を上限とする
  //   amounts.js の getMonthlyBilledStatus_ が治癒後別負傷 [B] を正しく判定するよう修正済みのため、
  //   transferData 側も validInitCount に合わせて初検料算定件数を反映する。
  //   [A] 施術継続中 Mixed: Math.min(rawInitCount, 1) と等価
  //   [B] 治癒後別負傷:     Math.min(rawInitCount, 2) で 2 件を反映
  // ★ reCount: 有効初検数(validInitCount)を上限とする（[A]=1 / [B]=2）
  return {
    initCount: Math.min(rawInitCount, validInitCount),
    reCount:   Math.min(rawReCount, validInitCount),
  };
}

function V3TR_buildCaseMoneyBlock_(agg) {
  const koryoSum = Math.round(agg.koryoSum);
  const coldSum  = Math.round(agg.coldSum);
  const warmSum  = Math.round(agg.warmSum);
  const elecSum  = Math.round(agg.elecSum);

  const koryoCount = agg.koryoCount || 0;
  const koryoUnit = (koryoCount > 0) ? Math.round(koryoSum / koryoCount) : 0;

  // 施療料（初検日の部位基本料）を全部位分合算してcaseTotalに算入する。
  // V3TR_aggregateDetailMonthly_ で kubun="初検" のbaseはkoryoSumに含めず
  // ptgt.shoryoFee に個別保存されているため、ここで明示的に合算する。
  const shoryoSum = Math.round(
    Object.values(agg.parts || {}).reduce(function(s, p) { return s + (p.shoryoFee || 0); }, 0)
  );

  const caseTotal = shoryoSum + koryoSum + coldSum + warmSum + elecSum;

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

  // RC-1整合: 転記シートに旧 case2 行が残っていても
  // caseKey が空 かつ case計=0 ならデータなしとして null 扱いにする。
  // exportTransferJson_（B案経路）と同じ判定条件に統一する。
  if (row2 && !String(row2["caseKey"] || "").trim() && Number(row2["case計"] || 0) === 0) {
    row2 = null;
  }

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

  /** テンプレ形式 {cell, tmpl} に数値を埋め込んで書き込む */
  function putTmpl(spec, val) {
    const n = Number(val);
    if (!isFinite(n) || n === 0) return;
    const text = spec.tmpl.replace("{amt}", n.toLocaleString());
    sh.getRange(spec.cell).setValue(text);
    count++;
  }

  /** 数値を1桁ずつ右詰めで配列セルに書き込む（合計欄用） */
  function putDigits(cellArr, val) {
    const n = Number(val);
    if (!isFinite(n) || n === 0) return;
    const digits = String(Math.round(n)).split("");
    // 右詰め: 配列末尾から埋める
    for (let i = cellArr.length - 1, d = digits.length - 1; i >= 0 && d >= 0; i--, d--) {
      sh.getRange(cellArr[i]).setValue(Number(digits[d]));
      count++;
    }
  }

  // ===== 保険者情報 =====
  // 保険者番号: 1桁ずつ（配列8セル）
  const insurerNo = String(row1["保険者番号"] || "").trim();
  if (insurerNo) {
    const insurerCells = CM.保険者番号;
    for (let i = 0; i < insurerNo.length && i < insurerCells.length; i++) {
      sh.getRange(insurerCells[i]).setValue(insurerNo.charAt(i));
      count++;
    }
  }
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
  // ★書き込み前に全 5 行を clearContent（fix: 2026-05-07）
  //   前回書き込み値が残りラベルなし重複が生じる問題を防ぐ
  // ★空行詰め: 名前が空のエントリを除外してから連続行に書き込む（fix: 2026-05-07）
  const injRows = CM.負傷名;

  // 先に全行クリア
  for (let ci = 0; ci < injRows.length; ci++) {
    const cm = injRows[ci];
    sh.getRange(cm.name).clearContent();
    [cm.injY, cm.injM, cm.injD,
     cm.iniY, cm.iniM, cm.iniD,
     cm.stY,  cm.stM,  cm.stD,
     cm.edY,  cm.edM,  cm.edD,
     cm.days].forEach(function(cell) {
      if (cell) sh.getRange(cell).clearContent();
    });
  }

  const injData    = V3TR_buildInjuryRows_(row1, row2);
  const validInj   = injData.filter(function(d) { return !!d.name; });

  for (let i = 0; i < validInj.length && i < injRows.length; i++) {
    const d = validInj[i];
    const m = injRows[i];

    // 行26（i=0）→（1）、行27（i=1）→（2）を先頭に付ける
    // 行28以降はテンプレートに (3)(4)(5) が既存のため番号不要
    var injName = d.name;
    if (i === 0) injName = "（1）" + injName;
    else if (i === 1) injName = "（2）" + injName;
    put(m.name, injName);
    V3TR_putDateYMD_(sh, m.injY, m.injM, m.injD, d.injuryDate);
    V3TR_putDateYMD_(sh, m.iniY, m.iniM, m.iniD, d.firstDate);
    V3TR_putDateYMD_(sh, m.stY, m.stM, m.stD, d.startDate);
    V3TR_putDateYMD_(sh, m.edY, m.edM, m.edD, d.endDate);
    putNum(m.days, d.days);
    count += V3TR_countDateWrites_(d);
  }

  // ===== 初検料・再検料 行33-34 =====
  // ★テンプレ形式: {cell, tmpl} → テキスト埋め込み
  putTmpl(CM.初検料, row1["初検料_月額"]);
  putTmpl(CM.初検時相談支援料, row1["初検時相談支援料_月額"]);
  putTmpl(CM.再検料, row1["再検料_月額"]);
  putTmpl(CM.基本3項目_計, row1["基本3項目_計"]);

  // ===== 施療料 行35 =====
  // ★施療料セルも {cell, no} 形式 → cell プロパティを使う
  // ※脱臼の整復料も「施療料」欄に記載する（申請書様式の規定）
  const shoryoData = V3TR_buildShoryoArray_(row1, row2);
  const shoryoCells = CM.施療料;
  for (let i = 0; i < shoryoData.length && i < shoryoCells.length; i++) {
    const amt = Number(shoryoData[i] || 0);
    if (amt > 0) {
      const label = "(" + shoryoCells[i].no + ")  " + amt.toLocaleString() + "円";
      sh.getRange(shoryoCells[i].cell).setValue(label);
      count++;
    }
  }
  // 施療料計 = case1 + case2
  const shoryoTotal = Number(row1["施療料_計"] || 0) + Number((row2 || {})["施療料_計"] || 0);
  putTmpl(CM.施療料_計, shoryoTotal);

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
  // ★1桁ずつ右詰め（配列6セル + "円"固定）
  putDigits(CM.合計, row1["当月合計"]);
  putDigits(CM.一部負担金, row1["窓口負担額"]);
  putDigits(CM.請求金額, row1["請求金額"]);

  // ===== U7 請求区分 行31 DH31 =====
  // "新規" or "継続" を書き込む。同月内治癒再発の両方○は将来対応。
  put(CM.請求区分, row1["請求区分"]);

  // ===== U5 本家区分 行8-13 =====
  // 判定ソース: 保険種別・続柄・生年月日・一部負担金割合・対象月
  // ★ 後期高齢者（保険種別=6 or 75歳以上）は制度上の記載方式未確認のため空欄（保留）
  // ★ 暫定ルール: docs/JREC-01_申請書様式運用メモ.md §4 U5 参照
  const HONKEKU_CIRCLE_MAP_ = {
    "本人": ["2", "②"], "六歳": ["4", "④"], "家族": ["6", "⑥"],
    "高一": ["8", "⑧"], "高7":  ["0", "⓪"],  // "高7" の "0" → Unicode U+24EA ⓪
  };
  const honkekuKey = V3TR_deriveHonkeku_(row1);
  if (honkekuKey && CM[honkekuKey]) {
    var hcCell = sh.getRange(CM[honkekuKey]);
    var hcVal  = String(hcCell.getValue() || "");
    var hm     = HONKEKU_CIRCLE_MAP_[honkekuKey];
    if (hm && hcVal.indexOf(hm[0]) !== -1) {
      hcCell.setValue(hcVal.replace(hm[0], hm[1]));
      count++;
    }
  }

  // ===== U6 給付割合 行8-13 =====
  // 片側丸付け: 対象数字1文字のみ置換（U5と同方式）
  // テンプレート実値: DP8='10・９' / DP11='８・７'（全角文字）
  // 割合=1→DP8('９'→'⑨') / 割合=2→DP11('８'→'⑧') / 割合=3→DP11('７'→'⑦')
  // 根拠: docs/JREC-01_申請書様式運用メモ.md §4 U6 参照
  const burden6 = Number(row1["一部負担金割合"]) || 0;
  var kyufuCharMap6 = {
    1: [CM.給付9割,  "９", "⑨"],
    2: [CM.給付8_7割, "８", "⑧"],
    3: [CM.給付8_7割, "７", "⑦"],
  };
  var kyufuEntry6 = kyufuCharMap6[burden6];
  if (kyufuEntry6 && kyufuEntry6[0]) {
    var kyufuRng6 = sh.getRange(kyufuEntry6[0]);
    var kyufuVal6 = String(kyufuRng6.getValue() || "");
    if (kyufuVal6.indexOf(kyufuEntry6[1]) !== -1) {
      kyufuRng6.setValue(kyufuVal6.replace(kyufuEntry6[1], kyufuEntry6[2]));
      count++;
    }
  }

  // ===== D4 負傷原因 行20 BR20 =====
  // 出力条件: 申請書記載部位の合計が3以上のとき
  // 根拠: 柔整療養費告示 別表第2 備考2「3部位目は所定料金の100分の60」
  // 部位スロット: スロット1=row1.部位1/スロット2=row1.部位2/スロット3=row2.部位1/スロット4=row2.部位2
  // 修正理由: 旧条件「row2["部位1_計"]>0」は case1(1部位)+case2(1部位)=2部位でもトリガーしていた（2026-04-20 修正）
  var _s1 = Number(row1["部位1_計"] || 0) > 0 || Number(row1["部位1_後療料_金額"] || 0) > 0;
  var _s2 = Number(row1["部位2_計"] || 0) > 0 || Number(row1["部位2_後療料_金額"] || 0) > 0;
  var _s3 = (row2 != null) && (Number(row2["部位1_計"] || 0) > 0 || Number(row2["部位1_後療料_金額"] || 0) > 0);
  var _s4 = (row2 != null) && (Number(row2["部位2_計"] || 0) > 0 || Number(row2["部位2_後療料_金額"] || 0) > 0);
  var _totalParts = (_s1 ? 1 : 0) + (_s2 ? 1 : 0) + (_s3 ? 1 : 0) + (_s4 ? 1 : 0);
  var part3HasData = _totalParts >= 3;
  if (part3HasData) {
    var d4Parts = [];
    function V3TR_buildInjuryText_(r) {
      var seg = [];
      var dt     = String(r["負傷の日時"] || "").trim();  // 日時→場所→状況
      var place  = String(r["負傷の場所"] || "").trim();
      var status = String(r["負傷の状況"] || "").trim();
      if (dt)     seg.push(dt);
      if (place)  seg.push(place);
      if (status) seg.push(status);
      return seg.join("\u3000"); // 全角スペース区切り
    }
    var d4T1 = V3TR_buildInjuryText_(row1);
    var d4T2 = row2 ? V3TR_buildInjuryText_(row2) : "";
    if (d4T1) d4Parts.push(d4T1);
    if (d4T2 && d4T2 !== d4T1) d4Parts.push(d4T2);
    // 1ケースにつき1行・行頭に（1）（2）番号付き・改行区切り
    var d4Lines = d4Parts.map(function(t, idx) { return "（" + (idx + 1) + "）" + t; });
    var d4Text = d4Lines.join("\n");
    if (d4Text) put(CM.負傷原因, d4Text);
  }

  // ===== 施術機関固定情報（設定シートから取得）=====
  // U1: 都道府県番号 → CI2 / U2: 施術機関コード → CZ2 / U4: 単独 → CT8
  // 下段登録記号番号 → CR51(左)/DK51(中)/DR51(右) 分割書込
  // ★ U2 は 登録記号番号 から先頭「協/契」のみ除去（ハイフン保持）した値を使う（暫定運用）
  const shSettingsForClinic = ss.getSheetByName(V3TR.CONFIG.sheetNames.settings);
  const clinicInfo = V3TR_loadClinicInfo_(shSettingsForClinic);
  put(CM.都道府県番号, clinicInfo.prefectureNo);
  put(CM.施術機関コード, V3TR_deriveClinicCode_(clinicInfo.torokuKigoNo));
  // U4 単独: テンプレート "1.単独" の "1" を "①" に置換
  if (clinicInfo.prefectureNo || clinicInfo.torokuKigoNo) {
    var tankeiCell = sh.getRange(CM.単独);
    var tankeiVal = tankeiCell.getValue();
    if (String(tankeiVal).indexOf("1") !== -1) {
      tankeiCell.setValue(String(tankeiVal).replace("1", "①"));
      count++;
    }
  }
  // 下段 登録記号番号 → CR51/DK51/DR51（分割書込）
  // CR49:DV50 はラベル行「登録記号番号」→ 書き込まない
  // 例: '契2804440-0-0' → 左='契2804440' / 中='0' / 右='0'
  var torokuFull = String(clinicInfo.torokuKigoNo || "").trim();
  if (torokuFull && CM.登録記号番号_左) {
    var torokuParts = torokuFull.split("-");
    put(CM.登録記号番号_左, torokuParts[0] || "");
    put(CM.登録記号番号_中, torokuParts.length > 1 ? torokuParts[1] : "");
    put(CM.登録記号番号_右, torokuParts.length > 2 ? torokuParts[2] : "");
  }

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

  // ★申請書は和暦年で記載する（appCellMapコメント参照）
  const era = V3TR_toWareki_(d);
  sh.getRange(cellY).setValue(era.year);
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
/**
 * @param {boolean} [skipBuild=false] true のとき build を省略（B案ループ内など build 済みの場合）
 */
function V3TR_exportTransferJson_(ss, patientId, ym, skipBuild) {
  // まず転記データを生成/更新（skipBuild=true の場合は呼び出し元で build 済みのため省略）
  if (!skipBuild) V3TR_buildTransferDataForMonth_(ss, patientId, ym);

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

  // RC-1整合: シートに旧 case2 行が残っていても
  // caseKey が空 かつ case計=0 ならデータなしとして null 扱いにする。
  // V3TR_buildTransferDataForMonth_ の guard（空行追加抑制）と条件を統一する。
  if (row2 && !String(row2["caseKey"] || "").trim() && Number(row2["case計"] || 0) === 0) {
    row2 = null;
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
      // 「対象月」は月フィールド（yyyy-MM）であり日付（yyyy-MM-dd）にしない
      if (key === "対象月") {
        out[key] = Utilities.formatDate(v, "Asia/Tokyo", "yyyy-MM");
      } else {
        out[key] = Utilities.formatDate(v, "Asia/Tokyo", "yyyy-MM-dd");
      }
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

/* =======================================================================
   一括バッチ出力（月指定 → 全患者NDJSON）
   ======================================================================= */

/**
 * メニュー：一括JSON出力（月指定）
 * 当月来院の全患者を抽出し、NDJSON形式でDriveに出力＋シートにバックアップ
 */
function V3TR_menuBatchExportJson() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt("一括JSON出力", "対象月（yyyy-MM）を入力してください（例：2026-02）", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var ym = (r1.getResponseText() || "").trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) return ui.alert("形式が違います。yyyy-MM で入力してください。");

  // 1. 当月来院の全患者IDを抽出
  var patientIds = V3TR_findPatientsForMonth_(ss, ym);
  if (patientIds.length === 0) {
    return ui.alert("対象月（" + ym + "）に来院記録のある患者が見つかりません。");
  }

  ui.alert("対象患者: " + patientIds.length + " 名\n" + patientIds.join(", ") + "\n\n処理を開始します。");

  // 2. 各患者のJSONを生成
  var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ssXXX");
  // 施術機関固定情報（設定シートから取得して meta に埋め込む）
  var shSettingsA = ss.getSheetByName(V3TR.CONFIG.sheetNames.settings);
  var clinicInfoA = V3TR_loadClinicInfo_(shSettingsA);
  var metaLine = JSON.stringify({
    _meta: true,
    schemaVersion: "3.0",
    generatedAt: now,
    month: ym,
    patientCount: patientIds.length,
    prefectureNo:       clinicInfoA.prefectureNo,
    torokuKigoNo:       clinicInfoA.torokuKigoNo,
    clinicName:         clinicInfoA.clinicName,
    clinicAddr:         clinicInfoA.clinicAddr,
    clinicPractitioner: clinicInfoA.clinicPractitioner,
  });

  var ndjsonLines = [metaLine];
  var sheetRows = [];  // バックアップ用

  for (var i = 0; i < patientIds.length; i++) {
    var pid = patientIds[i];
    ss.toast("処理中: " + pid + " (" + (i + 1) + "/" + patientIds.length + ")", "一括JSON出力", 3);

    try {
      // 転記データ生成（upsert）
      V3TR_buildTransferDataForMonth_(ss, pid, ym);

      // JSON生成（既存関数を利用）
      var jsonStr = V3TR_exportTransferJson_(ss, pid, ym);
      var parsed = JSON.parse(jsonStr);

      // ★Layer 2 安全フィルタ: 保険請求額=0 の患者は申請対象外として除外
      var _eL2c1 = parsed.case1 ? Number(parsed.case1["請求金額"] || 0) : 0;
      var _eL2c2 = parsed.case2 ? Number(parsed.case2["請求金額"] || 0) : 0;
      if (_eL2c1 === 0 && _eL2c2 === 0) {
        Logger.log("申請対象外スキップ（保険請求額=0）: " + pid);
        ss.toast("スキップ（申請対象外）: " + pid, "一括JSON出力", 3);
        sheetRows.push([pid, "SKIP: 保険請求額=0円（申請対象外）"]);
        continue;
      }

      // NDJSON行: { patientId, case1, case2, visitDays }
      var patientLine = JSON.stringify({
        patientId: pid,
        case1: parsed.case1,
        case2: parsed.case2,
        visitDays: parsed.visitDays || []
      });
      ndjsonLines.push(patientLine);
      sheetRows.push([pid, patientLine]);
    } catch (e) {
      Logger.log("患者 " + pid + " でエラー: " + e.message);
      ss.toast("エラー: " + pid + " - " + e.message, "一括JSON出力", 5);
      // エラーの患者はスキップして続行
      sheetRows.push([pid, "ERROR: " + e.message]);
    }
  }

  // ① patientCount後補正: スキップ発生時の validate_batch_safe 不一致を防ぐ
  var actualMetaA = JSON.parse(metaLine);
  actualMetaA.patientCount = ndjsonLines.length - 1;  // meta行を除いた実患者数
  ndjsonLines[0] = JSON.stringify(actualMetaA);

  // 3. Drive に NDJSON ファイルを出力
  var ndjsonContent = ndjsonLines.join("\n");
  var fileName = "transfer_batch_" + ym + ".ndjson";
  var folder = V3TR_getApplicationOutputFolder_(ss, ym);
  var archiveFolder = V3TR_getArchiveOutputFolder_(ss, ym);
  V3TR_archiveExistingFilesByExactName_(folder, archiveFolder, fileName);
  var file = folder.createFile(fileName, ndjsonContent, "application/x-ndjson");

  // 4. _JSON出力シートにバックアップ
  V3TR_writeSheetBackup_(ss, ym, now, sheetRows);

  // 5. 完了通知
  ui.alert(
    "一括JSON出力 完了\n\n" +
    "対象月: " + ym + "\n" +
    "患者数: " + patientIds.length + " 名\n" +
    "NDJSON行数: " + ndjsonLines.length + " 行（メタ1行＋患者" + (ndjsonLines.length - 1) + "行）\n\n" +
    "Drive出力: " + file.getName() + "\n" +
    "フォルダ: " + folder.getName() + "\n" +
    "シートバックアップ: _JSON出力"
  );
}

/**
 * 当月来院の全患者IDを抽出（施術明細＋来院ヘッダ両方をスキャン）
 * @param {Spreadsheet} ss
 * @param {string} ym - yyyy-MM
 * @return {string[]} ソート済み患者IDリスト
 */
function V3TR_findPatientsForMonth_(ss, ym) {
  var month = V3TR_parseYM_(ym);
  var patientSet = {};  // Object as Set（GAS互換）

  // --- 施術明細からスキャン（メインソース） ---
  var shDetail = ss.getSheetByName(V3TR.CONFIG.sheetNames.detail);
  if (shDetail && shDetail.getLastRow() >= 2) {
    var dMap = V3TR_buildHeaderMap_(shDetail);
    var cPid = dMap[V3TR.CONFIG.detailCols.patientId];
    var cDt = dMap[V3TR.CONFIG.detailCols.treatDate];
    if (cPid !== undefined && cDt !== undefined) {
      var dv = shDetail.getDataRange().getValues();
      for (var r = 1; r < dv.length; r++) {
        var dt = dv[r][cDt];
        if (V3TR_inRange_(dt, month.start, month.end)) {
          var pid = String(dv[r][cPid] || "").trim();
          if (pid) patientSet[pid] = true;
        }
      }
    }
  }

  // --- 来院ヘッダからスキャン（補助ソース、明細欠けの患者も拾う） ---
  // ★Layer 1 安全フィルタ: 会計区分=「自費のみ」の行は保険申請対象外なのでスキップ。
  //   自費のみ来院は施術明細に記録されないため、来院ヘッダ経由でのみ申請リストに混入するリスクがある。
  //   安全ルール: 保険申請対象 = 会計区分 ∈ {保険のみ, 保険+自費, ""(旧データ)} の行のみ通す。
  //   列が存在しない場合（旧データ等）は安全方向でスキップせず通す。
  var shHeader = ss.getSheetByName("来院ヘッダ");
  if (shHeader && shHeader.getLastRow() >= 2) {
    var hMap = V3TR_buildHeaderMap_(shHeader);
    var hPid = hMap["患者ID"];
    var hDt = hMap["施術日"];
    var hAcct = hMap["会計区分"];  // Layer 1: 申請対象フィルタ用
    if (hPid !== undefined && hDt !== undefined) {
      var hv = shHeader.getDataRange().getValues();
      for (var r = 1; r < hv.length; r++) {
        var dt = hv[r][hDt];
        if (V3TR_inRange_(dt, month.start, month.end)) {
          // Layer 1: 会計区分=「自費のみ」の行は申請リストに含めない
          if (hAcct !== undefined && String(hv[r][hAcct] || "").trim() === "自費のみ") continue;
          var pid = String(hv[r][hPid] || "").trim();
          if (pid) patientSet[pid] = true;
        }
      }
    }
  }

  return Object.keys(patientSet).sort();
}

/**
 * 出力先Driveフォルダを取得
 * 優先順: 設定シートの「出力フォルダID」→ スプレッドシート親フォルダ → ルート
 */
function V3TR_getOutputFolder_(ss) {
  if (typeof V3OUT !== 'undefined' && V3OUT.getOrCreateMonthlyOutputRootFolder_) {
    return V3OUT.getOrCreateMonthlyOutputRootFolder_(ss, '');
  }

  // 設定シートから出力フォルダIDを取得
  var shSettings = ss.getSheetByName(V3TR.CONFIG.sheetNames.settings);
  if (shSettings) {
    var sv = shSettings.getDataRange().getValues();
    for (var r = 0; r < sv.length; r++) {
      if (String(sv[r][0] || "").trim() === V3TR.CONFIG.setKeys.outputFolderId) {
        var folderId = String(sv[r][1] || "").trim();
        if (folderId) {
          try {
            return DriveApp.getFolderById(folderId);
          } catch (e) {
            Logger.log("出力フォルダID無効: " + folderId + " - " + e.message);
          }
        }
        break;
      }
    }
  }

  // フォールバック: スプレッドシートの親フォルダ
  try {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) {
      return parents.next();
    }
  } catch (e) {
    Logger.log("親フォルダ取得失敗: " + e.message);
  }

  // 最終フォールバック: ルート
  return DriveApp.getRootFolder();
}

/**
 * 申請書の月次保存先フォルダを返す。
 * 新ルール: JREC-01_月次出力/YYYY-MM/01_申請書/
 */
function V3TR_getApplicationOutputFolder_(ss, ym) {
  if (typeof V3OUT !== 'undefined' && V3OUT.getOrCreateDocTypeFolder_) {
    return V3OUT.getOrCreateDocTypeFolder_(ss, ym, 'application', '');
  }
  return V3TR_getOrCreateMonthFolder_(V3TR_getOutputFolder_(ss), ym);
}

/**
 * 月次再生成の旧版退避先フォルダを返す。
 * 新ルール: JREC-01_月次出力/YYYY-MM/90_再生成旧版/
 */
function V3TR_getArchiveOutputFolder_(ss, ym) {
  if (typeof V3OUT !== 'undefined' && V3OUT.getOrCreateArchiveFolder_) {
    return V3OUT.getOrCreateArchiveFolder_(ss, ym, '');
  }
  return V3TR_getOrCreateMonthFolder_(V3TR_getOutputFolder_(ss), ym);
}

function V3TR_archiveExistingFilesByExactName_(sourceFolder, archiveFolder, fileName) {
  if (typeof V3OUT !== 'undefined' && V3OUT.archiveFilesByExactName_) {
    return V3OUT.archiveFilesByExactName_(sourceFolder, archiveFolder, fileName);
  }

  var existing = sourceFolder.getFilesByName(fileName);
  var movedCount = 0;
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
    movedCount++;
  }
  return movedCount;
}

function V3TR_archiveExistingApplicationFiles_(sourceFolder, archiveFolder, patientId, ym) {
  var prefix = "申請書_" + String(patientId || "").trim() + "_" + String(ym || "").trim() + "_";
  if (typeof V3OUT !== 'undefined' && V3OUT.archiveFilesByPrefix_) {
    return V3OUT.archiveFilesByPrefix_(sourceFolder, archiveFolder, prefix);
  }
  return 0;
}

/**
 * ===== 申請書生成 B案メニュー =====
 * スプレッドシートから1クリックで申請書 xlsx を生成して Drive に保存する。
 *
 * 処理フロー:
 *   1. HTMLダイアログで患者ID（省略可）と対象月を入力
 *   2. NDJSON 生成（既存ロジック流用）
 *   3. Cloud Run POST /generate（X-Secret-Key 認証）
 *   4. レスポンス base64 → Drive の月別フォルダに xlsx 保存
 *   5. _申請書生成ログ シートへ記録
 *   6. 完了メッセージをダイアログ内に表示
 *
 * 設定（スクリプトプロパティ）:
 *   APPGEN_ENDPOINT  - Cloud Run エンドポイント URL
 *   APPGEN_SECRET    - X-Secret-Key の値
 */
function V3TR_menuGenerateApplication_B() {
  var defaultYm = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
  var html =
    '<!DOCTYPE html><style>' +
    'body{font-family:sans-serif;font-size:13px;margin:16px;line-height:1.5}' +
    'label{display:block;margin-top:10px;font-weight:bold}' +
    'input{width:100%;box-sizing:border-box;margin-top:4px;padding:5px;font-size:13px}' +
    '.note{color:#888;font-size:11px;margin-top:2px}' +
    '.buttons{margin-top:18px}' +
    'button{padding:7px 18px;margin-right:8px;cursor:pointer;font-size:13px;border-radius:4px}' +
    '.primary{background:#1a73e8;color:#fff;border:none}' +
    '.secondary{background:#fff;border:1px solid #ccc}' +
    '#msg{margin-top:10px;color:red;white-space:pre-wrap;font-size:12px}' +
    '#result{display:none;margin-top:10px;white-space:pre-wrap;font-size:12px}' +
    '</style>' +
    '<form id="frm">' +
    '<label>患者ID <span style="font-weight:normal;color:#888">（省略 = 対象月の全患者）</span></label>' +
    '<input id="pid" type="text" placeholder="例: P001">' +
    '<label>対象月</label>' +
    '<input id="ym" type="text" value="' + defaultYm + '" placeholder="yyyy-MM">' +
    '<div class="note">例: ' + defaultYm + '</div>' +
    '<div class="buttons">' +
    '<button class="primary" type="button" onclick="run()">生成</button>' +
    '<button class="secondary" type="button" onclick="google.script.host.close()">キャンセル</button>' +
    '</div></form>' +
    '<div id="msg"></div>' +
    '<div id="result"></div>' +
    '<script>' +
    'function run(){' +
    '  var pid=document.getElementById("pid").value.trim();' +
    '  var ym=document.getElementById("ym").value.trim();' +
    '  if(!/^\\d{4}-\\d{2}$/.test(ym)){document.getElementById("msg").textContent="対象月は yyyy-MM 形式で入力してください。";return;}' +
    '  document.getElementById("msg").textContent="処理中… 別ダイアログが表示された場合はそちらを操作してください。";' +
    '  document.querySelectorAll("button").forEach(function(b){b.disabled=true;});' +
    '  google.script.run' +
    '    .withSuccessHandler(function(res){' +
    '      document.getElementById("frm").style.display="none";' +
    '      document.getElementById("msg").textContent="";' +
    '      document.getElementById("result").style.display="block";' +
    '      document.getElementById("result").textContent=res;' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.querySelectorAll("button").forEach(function(b){b.disabled=false;});' +
    '      document.getElementById("msg").textContent="\u26a0 "+e.message;' +
    '    })' +
    '    .V3TR_runGenerateApplicationDialog(pid,ym);' +
    '}' +
    '<\/script>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(380).setHeight(240),
    "申請書生成（B案）"
  );
}

function V3TR_runGenerateApplicationDialog(patientId, ym) {
  var pid = String(patientId || "").trim();
  var ss  = SpreadsheetApp.getActive();
  var patientIds;
  if (pid) {
    patientIds = [pid];
  } else {
    patientIds = V3TR_findPatientsForMonth_(ss, ym);
    if (patientIds.length === 0) {
      throw new Error("対象月（" + ym + "）に来院記録のある患者が見つかりません。");
    }
  }
  return V3TR_generateApplicationBCore_(patientIds, ym);
}

function V3TR_generateApplicationBCore_(patientIds, ym) {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  // ===== 2. 設定取得 =====
  var props = PropertiesService.getScriptProperties();
  var endpoint = props.getProperty("APPGEN_ENDPOINT") || "";
  var secretKey = props.getProperty("APPGEN_SECRET") || "";

  if (!endpoint) {
    throw new Error("スクリプトプロパティ「APPGEN_ENDPOINT」が未設定です。\n" +
      "スクリプトエディタ > プロジェクトの設定 > スクリプトプロパティ から設定してください。");
  }
  if (!secretKey) {
    throw new Error("スクリプトプロパティ「APPGEN_SECRET」が未設定です。");
  }

  // ===== 3. NDJSON 生成 =====
  ss.toast("NDJSON 生成中... (" + patientIds.length + " 名)", "申請書生成", 5);

  var genAt = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ssXXX");
  // 施術機関固定情報（設定シートから取得して meta に埋め込む）
  var shSettingsB = ss.getSheetByName(V3TR.CONFIG.sheetNames.settings);
  var clinicInfoB = V3TR_loadClinicInfo_(shSettingsB);
  var metaLine = JSON.stringify({
    _meta: true,
    schemaVersion: "3.0",
    generatedAt: genAt,
    month: ym,
    patientCount: patientIds.length,
    prefectureNo:       clinicInfoB.prefectureNo,
    torokuKigoNo:       clinicInfoB.torokuKigoNo,
    clinicName:         clinicInfoB.clinicName,
    clinicAddr:         clinicInfoB.clinicAddr,
    clinicPractitioner: clinicInfoB.clinicPractitioner,
  });
  var ndjsonLines = [metaLine];

  var skipPatients = [];
  for (var i = 0; i < patientIds.length; i++) {
    var pid = patientIds[i];
    try {
      // P1 二重build除去: build後に skipBuild=true で export を呼ぶ（1患者1回）
      V3TR_buildTransferDataForMonth_(ss, pid, ym);
      var jsonStr = V3TR_exportTransferJson_(ss, pid, ym, true);
      var parsed = JSON.parse(jsonStr);
      // ★Layer 2 安全フィルタ: 保険請求額=0 の患者は申請対象外として除外。
      //   Layer 1（会計区分フィルタ）をすり抜けた場合の安全網。
      //   claimPay=0 = 保険申請する金額がない = 申請書を生成してはならない。
      var _bL2c1 = parsed.case1 ? Number(parsed.case1["請求金額"] || 0) : 0;
      var _bL2c2 = parsed.case2 ? Number(parsed.case2["請求金額"] || 0) : 0;
      if (_bL2c1 === 0 && _bL2c2 === 0) {
        Logger.log("[B案] 申請対象外スキップ（保険請求額=0）: " + pid);
        skipPatients.push(pid + "（申請対象外: 保険請求額=0円）");
        continue;
      }
      ndjsonLines.push(JSON.stringify({
        patientId: pid,
        case1: parsed.case1,
        case2: parsed.case2,
        visitDays: parsed.visitDays || []
      }));
    } catch (e) {
      Logger.log("[B案] NDJSON生成エラー: " + pid + " - " + e.message);
      skipPatients.push(pid + "（" + e.message + "）");
    }
  }

  // ===== ②-A プリフライトバリデーション（Cloud Run POST 前）=====
  // NDJSON生成は成功したが内容に問題がある患者を検出し、人間に判断を委ねる
  var PREFLIGHT_REQUIRED_KEYS = ["患者ID", "対象月", "患者氏名", "当月合計", "窓口負担額", "請求金額"];
  var preflightErrors   = [];  // [{pid, reasons:[]}]  hard error → 除外対象
  var preflightWarnings = [];  // [{pid, warnings:[]}] warning   → 確認後続行可

  for (var pi = 1; pi < ndjsonLines.length; pi++) {
    var pLine;
    try { pLine = JSON.parse(ndjsonLines[pi]); } catch (pe) {
      preflightErrors.push({ pid: "(行" + (pi + 1) + ")", reasons: ["JSONパース失敗: " + pe.message] });
      continue;
    }
    var ppid = pLine.patientId || "(患者ID不明)";
    var reasons  = [];
    var warnings = [];
    var c1 = pLine.case1;
    if (!c1) {
      reasons.push("case1 が null（転記データ未生成の可能性）");
    } else {
      // --- hard error: 必須キー空 ---
      PREFLIGHT_REQUIRED_KEYS.forEach(function(k) {
        var v = c1[k];
        if (v === null || v === undefined || String(v).trim() === "") {
          reasons.push("必須項目「" + k + "」が空");
        }
      });
      // --- hard error: 対象月不一致 ---
      var c1Month = String(c1["対象月"] || "").trim();
      if (c1Month && c1Month !== ym) {
        reasons.push("対象月不一致（データ=" + c1Month + ", 実行月=" + ym + "）");
      }

      // --- warning: 0値・未確定値（当月合計 > 0 の場合のみ評価）---
      var totalAmt  = Number(c1["当月合計"]    || 0);
      var copayAmt  = Number(c1["窓口負担額"]   || 0);
      var claimAmt  = Number(c1["請求金額"]     || 0);
      var burdenDig = c1["一部負担金割合"];
      var burdenVal = (burdenDig === null || burdenDig === undefined) ? 0 : Number(burdenDig);
      if (totalAmt > 0) {
        if (!burdenDig || burdenVal === 0) {
          warnings.push("一部負担金割合が 0 または未設定（負担割合が不明な可能性）");
        }
        if (copayAmt === 0) {
          warnings.push("当月合計 " + totalAmt + " 円 なのに窓口負担額が 0（免除・後日確認の場合は無視可）");
        }
        if (claimAmt === 0) {
          warnings.push("当月合計 " + totalAmt + " 円 なのに請求金額が 0（要確認）");
        }
        if (copayAmt + claimAmt !== totalAmt) {
          warnings.push("窓口負担額(" + copayAmt + ") + 請求金額(" + claimAmt + ") = " +
            (copayAmt + claimAmt) + " が当月合計(" + totalAmt + ")と一致しない");
        }
      }
    }
    if (reasons.length  > 0) preflightErrors.push({   pid: ppid, reasons:  reasons  });
    if (warnings.length > 0) preflightWarnings.push({ pid: ppid, warnings: warnings });
  }

  // ===== ②-B プリフライト結果処理（hard error）=====
  if (preflightErrors.length > 0) {
    var pfLines = ["【確認】以下の患者にデータ上の問題が見つかりました。\n"];
    preflightErrors.forEach(function(e) {
      pfLines.push("  " + e.pid + ":");
      e.reasons.forEach(function(r) { pfLines.push("    ・" + r); });
    });
    pfLines.push("\nこの患者をスキップして他の患者のみ生成しますか？");
    pfLines.push("（「はい」: 問題患者を除外して続行 ／ 「いいえ」: 中断）");

    var pfResp = ui.alert("申請書生成（B案）", pfLines.join("\n"), ui.ButtonSet.YES_NO);
    if (pfResp !== ui.Button.YES) {
      throw new Error("申請書生成を中断しました。\n問題患者のデータを確認してから再実行してください。");
    }

    // 問題患者を NDJSON から除外
    var errorPids = {};
    preflightErrors.forEach(function(e) { errorPids[e.pid] = true; });
    var filteredLines = [ndjsonLines[0]];  // meta行は保持
    for (var fi = 1; fi < ndjsonLines.length; fi++) {
      var fl;
      try { fl = JSON.parse(ndjsonLines[fi]); } catch(_) { continue; }
      if (!errorPids[fl.patientId]) {
        filteredLines.push(ndjsonLines[fi]);
      } else {
        // skipPatients に転記（除外理由付き）
        var pfErr = preflightErrors.filter(function(e) { return e.pid === fl.patientId; })[0];
        var pfReason = pfErr ? pfErr.reasons.join(" / ") : "プリフライト除外";
        skipPatients.push(fl.patientId + "（プリフライト除外: " + pfReason + "）");
        Logger.log("[B案] プリフライト除外: " + fl.patientId + " - " + pfReason);
      }
    }
    ndjsonLines = filteredLines;

    if (ndjsonLines.length === 1) {
      throw new Error("問題患者を除外した結果、生成対象の患者がいなくなりました。\n処理を中断します。");
    }
  }

  // ===== ②-C プリフライト結果処理（warning）=====
  // hard error で除外された患者の warning は表示しない（もう対象外のため）
  var excludedByError = {};
  preflightErrors.forEach(function(e) { excludedByError[e.pid] = true; });
  var activeWarnings = preflightWarnings.filter(function(w) { return !excludedByError[w.pid]; });

  if (activeWarnings.length > 0) {
    var wfLines = ["【要確認】以下の患者に未確定の可能性がある値があります。\n"];
    activeWarnings.forEach(function(w) {
      wfLines.push("  " + w.pid + ":");
      w.warnings.forEach(function(msg) { wfLines.push("    ⚠️ " + msg); });
    });
    wfLines.push("\n※正当な値であれば無視して構いません。");
    wfLines.push("このまま生成を続けますか？");
    wfLines.push("（「はい」: このまま続行 ／ 「いいえ」: 中断して確認）");

    var wfResp = ui.alert("申請書生成（B案）— 要確認", wfLines.join("\n"), ui.ButtonSet.YES_NO);
    if (wfResp !== ui.Button.YES) {
      throw new Error("申請書生成を中断しました。\n値を確認してから再実行してください。");
    }
    // 続行（患者除外なし）
    Logger.log("[B案] warning を確認のうえ続行: " + activeWarnings.map(function(w) { return w.pid; }).join(", "));
  }

  // ① patientCount後補正: スキップ・プリフライト除外後の実患者数で補正
  var actualMetaB = JSON.parse(ndjsonLines[0]);
  actualMetaB.patientCount = ndjsonLines.length - 1;  // meta行を除いた実患者数
  ndjsonLines[0] = JSON.stringify(actualMetaB);

  var ndjsonStr = ndjsonLines.join("\n");

  // ===== 4. Cloud Run へ POST =====
  ss.toast("Cloud Run に送信中...", "申請書生成", 10);

  var payload = JSON.stringify({ ndjson: ndjsonStr, month: ym });
  var options = {
    method: "post",
    contentType: "application/json",
    headers: { "X-Secret-Key": secretKey },
    payload: payload,
    muteHttpExceptions: true
  };

  var resp;
  try {
    resp = UrlFetchApp.fetch(endpoint + "/generate", options);
  } catch (e) {
    throw new Error("Cloud Run への接続に失敗しました:\n" + e.message +
      "\n\nエンドポイント: " + endpoint);
  }

  var statusCode = resp.getResponseCode();
  var bodyText = resp.getContentText();

  if (statusCode !== 200) {
    Logger.log("[B案] HTTP " + statusCode + ": " + bodyText);
    throw new Error("申請書生成に失敗しました (HTTP " + statusCode + ")\n\n" + bodyText.slice(0, 300));
  }

  var result;
  try {
    result = JSON.parse(bodyText);
  } catch (e) {
    throw new Error("レスポンスの解析に失敗しました:\n" + bodyText.slice(0, 300));
  }

  // ===== 5. Drive に保存 =====
  ss.toast("Drive に保存中...", "申請書生成", 10);

  var monthFolder = V3TR_getApplicationOutputFolder_(ss, ym);
  var archiveFolder = V3TR_getArchiveOutputFolder_(ss, ym);
  var timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "HHmmss");

  var savedFiles = [];
  var errorPatients = [];

  var patients = result.patients || [];
  for (var j = 0; j < patients.length; j++) {
    var p = patients[j];
    if (p.error || !p.content) {
      errorPatients.push(p.patientId + "（" + (p.error || "content なし") + "）");
      continue;
    }
    try {
      V3TR_archiveExistingApplicationFiles_(monthFolder, archiveFolder, p.patientId, ym);
      var fileName = "申請書_" + p.patientId + "_" + ym + "_" + timestamp + ".xlsx";
      var xlsxBlob = Utilities.newBlob(
        Utilities.base64Decode(p.content),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileName
      );
      var savedFile = monthFolder.createFile(xlsxBlob);
      savedFiles.push({ patientId: p.patientId, fileName: fileName, url: savedFile.getUrl(), warnings: p.warnings || [] });
    } catch (e) {
      errorPatients.push(p.patientId + "（Drive保存エラー: " + e.message + "）");
    }
  }

  // ===== 6. ログ記録 =====
  V3TR_writeGenerationLog_(ss, ym, genAt, savedFiles, errorPatients.concat(skipPatients));

  // ===== 7. 完了（結果を返す）=====
  var lines = [
    "【申請書生成完了】",
    "",
    "対象月: " + ym,
    "保存: " + savedFiles.length + " 件",
    "エラー: " + (errorPatients.length + skipPatients.length) + " 件",
    "",
    "保存先フォルダ: " + monthFolder.getUrl(),
  ];

  if (savedFiles.length > 0) {
    lines.push("");
    lines.push("生成ファイル:");
    savedFiles.forEach(function(f) {
      var warn = f.warnings.length > 0 ? " ⚠️要確認" : "";
      lines.push("  " + f.patientId + warn);
    });
  }
  if (errorPatients.length + skipPatients.length > 0) {
    lines.push("");
    lines.push("エラー患者:");
    errorPatients.concat(skipPatients).forEach(function(e) {
      lines.push("  " + e);
    });
  }

  return lines.join("\n");
}


/**
 * Web UI から呼ぶ B案申請書生成 wrapper（WEB-4A）
 * google.script.run から呼ばれる。APPGEN_SECRET/ENDPOINT は ScriptProperties から読む。
 * UI ダイアログなし。プリフライト hard error は JSON error として返す。
 * @param {string} patientId 患者ID
 * @param {string} ym 対象月（yyyy-MM）
 * @return {{ok,patientId,ym,fileName,fileUrl,message}|{ok,errorCode,message}}
 */
function generateClaimApplicationBFromWeb_V3(patientId, ym) {
  var pid   = String(patientId || "").trim();
  var ymStr = String(ym || "").trim();

  if (!pid) {
    return { ok: false, errorCode: "INVALID_PATIENT_ID", message: "患者IDが未指定です。" };
  }
  if (!/^\d{4}-\d{2}$/.test(ymStr)) {
    return { ok: false, errorCode: "INVALID_YM", message: "対象月の形式が不正です（yyyy-MM）。" };
  }

  var props     = PropertiesService.getScriptProperties();
  var endpoint  = props.getProperty("APPGEN_ENDPOINT") || "";
  var secretKey = props.getProperty("APPGEN_SECRET")   || "";

  if (!endpoint || !secretKey) {
    return { ok: false, errorCode: "APPGEN_CONFIG_MISSING", message: "申請書生成設定が不足しています。管理者に連絡してください。" };
  }

  var ss = SpreadsheetApp.getActive();

  // --- NDJSON 生成 ---
  var genAt        = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ssXXX");
  var shSettingsB  = ss.getSheetByName(V3TR.CONFIG.sheetNames.settings);
  var clinicInfoB  = V3TR_loadClinicInfo_(shSettingsB);
  var metaLine     = JSON.stringify({
    _meta: true, schemaVersion: "3.0", generatedAt: genAt, month: ymStr,
    patientCount: 1,
    prefectureNo:       clinicInfoB.prefectureNo,
    torokuKigoNo:       clinicInfoB.torokuKigoNo,
    clinicName:         clinicInfoB.clinicName,
    clinicAddr:         clinicInfoB.clinicAddr,
    clinicPractitioner: clinicInfoB.clinicPractitioner,
  });

  try { V3TR_buildTransferDataForMonth_(ss, pid, ymStr); } catch (e) {
    Logger.log("[WEB-4A] buildTransferData error pid=" + pid + ": " + e.message);
    return { ok: false, errorCode: "BUILD_FAILED", message: "転記データ生成に失敗しました。月次データを確認してください。" };
  }

  var jsonStr;
  try { jsonStr = V3TR_exportTransferJson_(ss, pid, ymStr, true); } catch (e) {
    Logger.log("[WEB-4A] exportTransferJson error pid=" + pid + ": " + e.message);
    return { ok: false, errorCode: "EXPORT_FAILED", message: "転記データのエクスポートに失敗しました。" };
  }

  var parsed;
  try { parsed = JSON.parse(jsonStr); } catch (e) {
    return { ok: false, errorCode: "PARSE_FAILED", message: "転記データの解析に失敗しました。" };
  }

  // Layer 2 安全フィルタ（保険請求額=0 は申請対象外）
  var claimC1 = parsed.case1 ? Number(parsed.case1["請求金額"] || 0) : 0;
  var claimC2 = parsed.case2 ? Number(parsed.case2["請求金額"] || 0) : 0;
  if (claimC1 === 0 && claimC2 === 0) {
    Logger.log("[WEB-4A] ZERO_CLAIM pid=" + pid);
    return { ok: false, errorCode: "ZERO_CLAIM", message: "保険請求額が0円のため申請書を生成できません。" };
  }

  // プリフライト（hard error のみ — warning は自動続行）
  var PREFLIGHT_REQUIRED_KEYS = ["患者ID", "対象月", "患者氏名", "当月合計", "窓口負担額", "請求金額"];
  var c1 = parsed.case1;
  if (!c1) {
    return { ok: false, errorCode: "PREFLIGHT_FAILED", message: "転記データが存在しません。転記データ生成を先に実行してください。" };
  }
  var missingKeys = PREFLIGHT_REQUIRED_KEYS.filter(function(k) {
    var v = c1[k];
    return v === null || v === undefined || String(v).trim() === "";
  });
  if (missingKeys.length > 0) {
    Logger.log("[WEB-4A] preflight missing keys pid=" + pid + ": " + missingKeys.join(", "));
    return { ok: false, errorCode: "PREFLIGHT_FAILED", message: "転記データに必須項目が不足しています（" + missingKeys.join(", ") + "）。スプレッドシートで確認してください。" };
  }
  var c1Month = String(c1["対象月"] || "").trim();
  if (c1Month && c1Month !== ymStr) {
    Logger.log("[WEB-4A] preflight month mismatch pid=" + pid + " data=" + c1Month + " exec=" + ymStr);
    return { ok: false, errorCode: "MONTH_MISMATCH", message: "転記データの対象月（" + c1Month + "）と実行月（" + ymStr + "）が一致しません。" };
  }

  var patientLine = JSON.stringify({ patientId: pid, case1: parsed.case1, case2: parsed.case2, visitDays: parsed.visitDays || [] });
  var ndjsonStr   = [metaLine, patientLine].join("\n");

  // --- Cloud Run POST ---
  var options = {
    method: "post",
    contentType: "application/json",
    headers: { "X-Secret-Key": secretKey },
    payload: JSON.stringify({ ndjson: ndjsonStr, month: ymStr }),
    muteHttpExceptions: true
  };

  var resp;
  try { resp = UrlFetchApp.fetch(endpoint + "/generate", options); } catch (e) {
    Logger.log("[WEB-4A] Cloud Run fetch error: " + e.message);
    return { ok: false, errorCode: "CLOUDRUN_UNREACHABLE", message: "申請書生成サーバーへの接続に失敗しました。" };
  }

  var statusCode = resp.getResponseCode();
  if (statusCode !== 200) {
    Logger.log("[WEB-4A] Cloud Run HTTP " + statusCode + ": " + resp.getContentText().slice(0, 200));
    return { ok: false, errorCode: "CLOUDRUN_ERROR", message: "申請書生成サーバーがエラーを返しました (HTTP " + statusCode + ")。" };
  }

  var result;
  try { result = JSON.parse(resp.getContentText()); } catch (e) {
    return { ok: false, errorCode: "RESPONSE_PARSE_FAILED", message: "サーバーレスポンスの解析に失敗しました。" };
  }

  // --- Drive 保存 ---
  var monthFolder;
  try { monthFolder = V3TR_getApplicationOutputFolder_(ss, ymStr); } catch (e) {
    return { ok: false, errorCode: "DRIVE_FOLDER_FAILED", message: "Drive出力フォルダの取得に失敗しました。" };
  }
  var archiveFolder = V3TR_getArchiveOutputFolder_(ss, ymStr);
  var timestamp     = Utilities.formatDate(new Date(), "Asia/Tokyo", "HHmmss");

  var patients = result.patients || [];
  if (patients.length === 0) {
    return { ok: false, errorCode: "NO_OUTPUT", message: "申請書が生成されませんでした（レスポンスが空）。" };
  }
  var p = patients[0];
  if (p.error || !p.content) {
    Logger.log("[WEB-4A] patient generation error: " + (p.error || "content なし"));
    return { ok: false, errorCode: "GENERATION_ERROR", message: "申請書生成でエラーが発生しました。設定または月次データを確認してください。" };
  }

  var savedFile, fileName;
  try {
    V3TR_archiveExistingApplicationFiles_(monthFolder, archiveFolder, pid, ymStr);
    fileName = "申請書_" + pid + "_" + ymStr + "_" + timestamp + ".xlsx";
    var xlsxBlob = Utilities.newBlob(
      Utilities.base64Decode(p.content),
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName
    );
    savedFile = monthFolder.createFile(xlsxBlob);
  } catch (e) {
    Logger.log("[WEB-4A] Drive save error: " + e.message);
    return { ok: false, errorCode: "DRIVE_SAVE_FAILED", message: "Driveへの保存に失敗しました。" };
  }

  // --- ログ記録（失敗しても生成成功として返す）---
  try {
    V3TR_writeGenerationLog_(ss, ymStr, genAt,
      [{ patientId: pid, fileName: fileName, url: savedFile.getUrl(), warnings: p.warnings || [] }], []);
  } catch (e) {
    Logger.log("[WEB-4A] ログ記録失敗（生成は成功）: " + e.message);
  }

  return {
    ok:        true,
    patientId: pid,
    ym:        ymStr,
    fileName:  fileName,
    fileUrl:   savedFile.getUrl(),
    message:   "申請書Excelを生成しました。Driveで確認してください。"
  };
}


/**
 * 月別フォルダを取得または作成する
 * 例: output/2026-03/
 */
function V3TR_getOrCreateMonthFolder_(parentFolder, ym) {
  var folders = parentFolder.getFoldersByName(ym);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(ym);
}


/**
 * _申請書生成ログ シートに生成結果を追記する
 */
function V3TR_writeGenerationLog_(ss, ym, generatedAt, savedFiles, errorList) {
  var LOG_SHEET = "_申請書生成ログ";
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.getRange(1, 1, 1, 8).setValues([[
      "実行日時", "対象月", "患者ID", "ファイル名", "Drive URL",
      "warnings", "ステータス", "エラー詳細"
    ]]);
    sh.setFrozenRows(1);
  }

  var rows = [];
  savedFiles.forEach(function(f) {
    rows.push([
      generatedAt, ym, f.patientId, f.fileName, f.url,
      f.warnings.join(" / "), "OK", ""
    ]);
  });
  errorList.forEach(function(e) {
    // "patientId（理由）" 形式のパース
    var match = e.match(/^([^（]+)（(.+)）$/);
    var pid = match ? match[1] : e;
    var detail = match ? match[2] : "";
    rows.push([
      generatedAt, ym, pid, "", "", "", "ERROR", detail
    ]);
  });

  if (rows.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  }
}


/**
 * _JSON出力シートにバックアップ書き込み
 */
function V3TR_writeSheetBackup_(ss, ym, generatedAt, patientRows) {
  var shJson = ss.getSheetByName("_JSON出力");
  if (!shJson) shJson = ss.insertSheet("_JSON出力");
  shJson.clear();

  // メタデータ
  var meta = [
    ["schemaVersion", "3.0"],
    ["generatedAt", generatedAt],
    ["month", ym],
    ["patientCount", patientRows.length],
    [],  // 空行
  ];

  // ヘッダー＋データ
  var header = ["patientId", "json"];
  var allRows = meta.concat([header]).concat(patientRows);

  if (allRows.length > 0) {
    // 各行を2列に正規化
    var normalized = allRows.map(function(row) {
      return [row[0] || "", row[1] || ""];
    });
    shJson.getRange(1, 1, normalized.length, 2).setValues(normalized);
  }
}
