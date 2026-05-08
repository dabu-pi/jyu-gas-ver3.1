/****************************************************
 * 柔整 Ver3.1 金額フェーズ（統合版対応）
 *
 * ★役割
 * (A) 共有ユーティリティ関数（core側からも呼ばれる）
 *     - loadSettings_V3_      … 設定シート読み込み
 *     - loadBurdenRatio_V3_   … 患者マスタから負担割合取得
 *     - calcBaseFee_V3_       … 区分×傷病種別→基本料
 *     - detectInjuryType_V3_  … 傷病名→打撲/捻挫/挫傷
 *     - roundToUnit_V3_       … 端数処理（四捨五入）
 *     - asDate_V3_ / diffDays_V3_ / pickInjuryDate_V3_
 *
 * (B) 施術明細シートベースの手動再計算
 *     - menuRecalcAmounts_V3  … メニュー実行
 *     - recalcAmountsByVisitKey_V3_ … 実処理
 *
 * ★注意
 * - core側の SHEETS / HEADER_COLS / MASTER_COLS / CASE_COLS /
 *   buildHeaderColMap_ / ensureRequiredCols_ 等を利用する
 * - このファイルでは重複宣言しない
 *
 * ★visitKey形式（統合版）
 * - 患者ID_yyyy-MM-dd（旧"|"から"_"に変更済み）
 ****************************************************/

/** ===== 設定キー（設定シートA列）===== */
const AM_SET_KEYS = {
  initFee: "初検料",
  initSupport: "初検時相談支援料",
  reFee: "再検料",
  // 打撲・捻挫・挫傷
  shoryoDaboku: "施療料_打撲",
  shoryoNenZa: "施療料_捻挫",
  shoryoZasyo: "施療料_挫傷",
  koryoDaboku: "後療料_打撲",
  koryoNenZa: "後療料_捻挫",
  koryoZasyo: "後療料_挫傷",
  // 脱臼（§17.2）
  seifukuDakkyu: "整復料_脱臼",
  koryoDakkyu: "後療料_脱臼",
  // 骨折・不全骨折（§17.3）
  koryoKossetu: "後療料_骨折",
  koryoFuzenKossetu: "後療料_不全骨折",
  seifukuKossetuPrefix: "整復料_骨折_",  // 動的キー: "整復料_骨折_前腕" 等
  koteiPrefix: "固定料_",                 // 動的キー: "固定料_前腕" 等
  cold: "冷罨法",
  warm: "温罨法",
  electro: "電療",
  taiki: "待機_打撲捻挫挫傷",
  multiCoef3: "多部位_3部位目係数",
  roundUnit: "窓口端数単位",
  metalAddon: "金属副子等加算",  // §18.3 骨折・不全骨折・脱臼のみ算定
  exerciseAddon: "柔道整復運動後療料",  // 骨折・不全骨折・脱臼 dayDiff>=15
};

/** ===== 施術明細：列名（最終ヘッダー前提）===== */
const AM_DETAIL_COLS = {
  // key
  detailID: "detailID",
  visitKey: "visitKey",
  patientId: "患者ID",
  treatDate: "施術日",
  kubun: "区分",
  caseNo: "caseNo",
  caseKey: "caseKey",
  injuryDateFixed: "受傷日_確定",
  injuryDateInput: "受傷日(入力)",
  bui: "部位",
  byomei: "傷病",
  partOrder: "部位順位",
  coldChk: "冷",
  warmChk: "温",
  electroChk: "電",
  metalChk: "金属副子チェック",  // §18.3 Phase 1
  exerciseChk: "運動後療チェック",  // 柔道整復運動後療料 Phase 1

  // 出力（確定列）
  coefOut: "係数",
  baseOut: "基本料_確定",
  supportOut: "初検相談_確定",
  coldOut: "冷_確定",
  warmOut: "温_確定",
  electroOut: "電_確定",
  taikiOut: "待機_確定",
  metalOut: "金属副子_確定",    // §18.3 Phase 1
  exerciseOut: "運動後療_確定",  // 柔道整復運動後療料 Phase 1
  rowTotalOut: "行合計_確定",
};


/* =======================================================================
   (A) 共有ユーティリティ関数
   ======================================================================= */

/** 設定読み込み（A:キー B:値） */
function loadSettings_V3_(ss) {
  var sh = ss.getSheetByName(SHEETS.settings);
  var values = sh.getDataRange().getValues();
  var map = {};

  for (var r = 1; r < values.length; r++) {
    var key = String(values[r][0] || "").trim();
    if (!key) continue;
    var val = values[r][1];
    map[key] = (typeof val === "number") ? val : Number(String(val || "").trim() || 0);
  }

  return {
    initFee: Number(map[AM_SET_KEYS.initFee] || 0),
    initSupport: Number(map[AM_SET_KEYS.initSupport] || 0),
    reFee: Number(map[AM_SET_KEYS.reFee] || 0),
    // 打撲・捻挫・挫傷
    shoryoDaboku: Number(map[AM_SET_KEYS.shoryoDaboku] || 0),
    shoryoNenZa: Number(map[AM_SET_KEYS.shoryoNenZa] || 0),
    shoryoZasyo: Number(map[AM_SET_KEYS.shoryoZasyo] || 0),
    koryoDaboku: Number(map[AM_SET_KEYS.koryoDaboku] || 0),
    koryoNenZa: Number(map[AM_SET_KEYS.koryoNenZa] || 0),
    koryoZasyo: Number(map[AM_SET_KEYS.koryoZasyo] || 0),
    // 脱臼（§17.2）
    seifukuDakkyu: Number(map[AM_SET_KEYS.seifukuDakkyu] || 0),
    koryoDakkyu: Number(map[AM_SET_KEYS.koryoDakkyu] || 0),
    // 骨折・不全骨折（§17.3）— 後療料は固定値
    koryoKossetu: Number(map[AM_SET_KEYS.koryoKossetu] || 0),
    koryoFuzenKossetu: Number(map[AM_SET_KEYS.koryoFuzenKossetu] || 0),
    cold: Number(map[AM_SET_KEYS.cold] || 0),
    warm: Number(map[AM_SET_KEYS.warm] || 0),
    electro: Number(map[AM_SET_KEYS.electro] || 0),
    taiki: Number(map[AM_SET_KEYS.taiki] || 0),
    multiCoef3: Number(map[AM_SET_KEYS.multiCoef3] || 0.6),
    roundUnit: Number(map[AM_SET_KEYS.roundUnit] || 1),
    metalAddon: Number(map[AM_SET_KEYS.metalAddon] || 0),  // §18.3
    exerciseAddon: Number(map[AM_SET_KEYS.exerciseAddon] || 0),  // 柔道整復運動後療料
    // 動的キー参照用（部位別整復料/固定料）
    _rawMap: map,
  };
}

/** 患者マスタから負担割合取得（0.3 or 30 どちらでもOK） */
function loadBurdenRatio_V3_(masterSh, masterMap, patientId) {
  var values = masterSh.getDataRange().getValues();
  var pidCol0 = masterMap[MASTER_COLS.patientId] - 1;
  var bCol0 = masterMap[MASTER_COLS.burden] - 1;

  for (var r0 = 1; r0 < values.length; r0++) {
    if (String(values[r0][pidCol0] || "").trim() !== patientId) continue;
    var raw = values[r0][bCol0];
    var num = (typeof raw === "number") ? raw : Number(String(raw || "").trim());
    if (!isFinite(num)) return 0;
    return (num > 1) ? (num / 100) : num;
  }
  throw new Error("患者マスタに患者ID=" + patientId + "が見つかりません。");
}

/**
 * 部位名(bui)から設定シートキーの部位コードを返す。
 * 例: "右前腕" → "前腕", "左鎖骨" → "鎖骨", "右第3中手骨" → "指_趾"
 * @param {string} bui - 部位名（患者画面のA列入力値）
 * @return {string|null} 設定シートの部位別キーに使うサフィックス
 */
function mapBuiToSettingKey_V3_(bui) {
  if (!bui) return null;
  // 長い文字列を先にマッチ（前腕 vs 腕 等の誤マッチ防止）
  var rules = [
    { keywords: ["鎖骨"], key: "鎖骨" },
    { keywords: ["肋骨"], key: "肋骨" },
    { keywords: ["前腕", "橈骨", "尺骨"], key: "前腕" },
    { keywords: ["上腕"], key: "上腕" },
    { keywords: ["下腿", "脛骨", "腓骨"], key: "下腿" },
    { keywords: ["大腿"], key: "大腿" },
    { keywords: ["指", "趾", "中手骨", "中足骨", "基節骨", "末節骨"], key: "指_趾" },
  ];
  for (var i = 0; i < rules.length; i++) {
    for (var j = 0; j < rules[i].keywords.length; j++) {
      if (bui.indexOf(rules[i].keywords[j]) !== -1) return rules[i].key;
    }
  }
  return null;
}

/** 基本料（設定シート通り） §17参照
 * @param {Object} settings - loadSettings_V3_ の返り値
 * @param {string} kubun - "初検" / "再検" / "後療"
 * @param {string} injuryType - detectInjuryType_V3_ の返り値
 * @param {string} [bui] - 部位名（骨折・不全骨折の初検日のみ必要）
 */
function calcBaseFee_V3_(settings, kubun, injuryType, bui) {
  if (kubun === "初検") {
    // 骨折: 整復料（部位別）
    if (injuryType === "骨折") {
      var buiKey = mapBuiToSettingKey_V3_(bui);
      if (buiKey && settings._rawMap) {
        var val = settings._rawMap[AM_SET_KEYS.seifukuKossetuPrefix + buiKey];
        if (val != null && isFinite(Number(val))) return Number(val);
      }
      return 0; // 部位不明 → 0（要確認フラグで補捉）
    }
    // 不全骨折: 固定料（部位別）
    if (injuryType === "不全骨折") {
      var buiKey2 = mapBuiToSettingKey_V3_(bui);
      if (buiKey2 && settings._rawMap) {
        var val2 = settings._rawMap[AM_SET_KEYS.koteiPrefix + buiKey2];
        if (val2 != null && isFinite(Number(val2))) return Number(val2);
      }
      return 0;
    }
    if (injuryType === "脱臼") return settings.seifukuDakkyu;  // 整復料（脱臼）
    if (injuryType === "打撲") return settings.shoryoDaboku;
    if (injuryType === "捻挫") return settings.shoryoNenZa;
    if (injuryType === "挫傷") return settings.shoryoZasyo;
    return 0;
  }
  if (kubun === "再検" || kubun === "後療") {
    if (injuryType === "骨折") return settings.koryoKossetu;         // 後療料（骨折）850円
    if (injuryType === "不全骨折") return settings.koryoFuzenKossetu; // 後療料（不全骨折）720円
    if (injuryType === "脱臼") return settings.koryoDakkyu;
    if (injuryType === "打撲") return settings.koryoDaboku;
    if (injuryType === "捻挫") return settings.koryoNenZa;
    if (injuryType === "挫傷") return settings.koryoZasyo;
    return 0;
  }
  return 0;
}

/** 傷病名→傷病種別 判別（基本料算定用）
 *  判定順序: 不全骨折→骨折→脱臼→打撲→捻挫→挫傷
 *  ※「不全骨折」は「骨折」を含むため先に判定する */
function detectInjuryType_V3_(byomei) {
  if (!byomei) return null;
  if (byomei.indexOf("不全骨折") !== -1) return "不全骨折";
  if (byomei.indexOf("骨折") !== -1) return "骨折";
  if (byomei.indexOf("脱臼") !== -1) return "脱臼";
  if (byomei.indexOf("打撲") !== -1) return "打撲";
  if (byomei.indexOf("捻挫") !== -1) return "捻挫";
  if (byomei.indexOf("挫傷") !== -1) return "挫傷";
  return null;
}

/** 端数処理（四捨五入） */
function roundToUnit_V3_(value, unit) {
  var u = Number(unit || 1);
  if (!isFinite(value)) return 0;
  if (!isFinite(u) || u <= 0) return Math.round(value);
  return Math.round(value / u) * u;
}

/** Date変換（Date/文字列→Date or null） */
function asDate_V3_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "string") {
    var s = v.trim();
    if (!s) return null;
    var d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

/** 日数差（受傷→施術） */
function diffDays_V3_(injury, treat) {
  if (!(injury instanceof Date) || !(treat instanceof Date)) return null;
  var a = new Date(injury.getFullYear(), injury.getMonth(), injury.getDate());
  var b = new Date(treat.getFullYear(), treat.getMonth(), treat.getDate());
  return Math.round((b.getTime() - a.getTime()) / (24 * 3600 * 1000));
}

/** 受傷日取得（確定列→入力列フォールバック） */
function pickInjuryDate_V3_(row, detailMap) {
  var cFixed = detailMap[AM_DETAIL_COLS.injuryDateFixed];
  if (cFixed) {
    var d = asDate_V3_(row[cFixed - 1]);
    if (d) return d;
  }
  var cIn = detailMap[AM_DETAIL_COLS.injuryDateInput];
  if (cIn) {
    var d2 = asDate_V3_(row[cIn - 1]);
    if (d2) return d2;
  }
  return null;
}


/* =======================================================================
   (B) 施術明細シートベースの手動再計算
   ======================================================================= */

/** メニュー実行（visitKeyを入力） */
function menuRecalcAmounts_V3() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  var res = ui.prompt(
    "金額再計算",
    "visitKey を入力してください（例：P0001_2026-02-15）",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var visitKey = (res.getResponseText() || "").trim();
  if (!visitKey) {
    ui.alert("visitKey が空です。");
    return;
  }

  var result = recalcAmountsByVisitKey_V3_(ss, visitKey);
  ui.alert(
    "完了：" + visitKey + "\n" +
    "明細行更新：" + result.updatedRows + "\n" +
    "来院合計：" + result.total + "\n" +
    "窓口負担：" + result.copay + "\n" +
    "保険請求：" + result.claim
  );
}

/**
 * 起算月からの経過月数を算出（§11 長期判定共通ユーティリティ）
 * 受傷日が16日以降なら翌月起算。Date でなければ -1 を返す。
 * @return {number} 経過月数（0以上）または -1（日付不正）
 */
function calcMonthsElapsed_V3_(injuryDate, treatDate) {
  if (!(injuryDate instanceof Date) || !(treatDate instanceof Date)) return -1;
  var sy = injuryDate.getFullYear(), sm = injuryDate.getMonth();
  if (injuryDate.getDate() >= 16) {
    sm++;
    if (sm > 11) { sm = 0; sy++; }
  }
  return (treatDate.getFullYear() - sy) * 12 + (treatDate.getMonth() - sm);
}

/**
 * 長期減額係数の算定（§11）
 * 骨折・不全骨折は対象外。受傷日の起算月から5か月超で75%。
 * 5か月超かつ月10回以上×5か月連続の場合は50%。
 * @param {number[]|null} monthlyVisitCounts - 起算月1〜5の月別来院回数配列（省略可）
 * @return {number} 1.0（減額なし）or 0.75（長期75%）or 0.50（長期50%）
 */
function calcLongTermCoef_V3_(injuryType, injuryDate, treatDate, monthlyVisitCounts) {
  // 骨折・不全骨折は長期減額の対象外
  if (injuryType === "骨折" || injuryType === "不全骨折") return 1.0;
  var monthsElapsed = calcMonthsElapsed_V3_(injuryDate, treatDate);
  if (monthsElapsed < 0) return 1.0;

  // 5か月超（6か月目以降）→ 50% or 75%
  if (monthsElapsed >= 5) {
    // 50%: 起算月1〜5が全て月10回以上（§11）
    if (Array.isArray(monthlyVisitCounts) && monthlyVisitCounts.length >= 5) {
      var allFrequent = true;
      for (var m = 0; m < 5; m++) {
        if (Number(monthlyVisitCounts[m] || 0) < 10) { allFrequent = false; break; }
      }
      if (allFrequent) return 0.50;
    }
    return 0.75;
  }
  return 1.0;
}

/**
 * 来院ヘッダデータから特定 caseKey の月別来院回数（起算月1〜5）を返す（§11 50%判定用）
 * 来院ヘッダ 1行 = 1 visit として集計。部位数は無関係。
 * @param {Array[][]} headerValues - 来院ヘッダ getDataRange().getValues()
 * @param {Object} headMap - buildHeaderColMap_ の結果（HEADER_COLS キー）
 * @param {string} patientId
 * @param {string} caseKey
 * @param {Date|null} injuryDate - 起算月計算用（null なら [0,0,0,0,0] を返す）
 * @return {number[]} 長さ5の配列 [月1来院数, 月2来院数, ..., 月5来院数]
 */
function buildMonthlyVisitCounts_V3_(headerValues, headMap, patientId, caseKey, injuryDate) {
  var empty = [0, 0, 0, 0, 0];
  if (!headerValues || !headMap || !patientId || !caseKey) return empty;
  if (!(injuryDate instanceof Date)) return empty;

  // 起算月（受傷日が16日以降なら翌月起算）
  var startYear  = injuryDate.getFullYear();
  var startMonth = injuryDate.getMonth(); // 0-indexed
  if (injuryDate.getDate() >= 16) {
    startMonth++;
    if (startMonth > 11) { startMonth = 0; startYear++; }
  }

  // 列インデックス（1-based → undefined なら早期リターン）
  var pidIdx = headMap[HEADER_COLS.patientId];
  var ckIdx  = headMap[HEADER_COLS.caseKey];
  var dtIdx  = headMap[HEADER_COLS.treatDate];
  if (!pidIdx || !ckIdx || !dtIdx) return empty;
  var pidCol0 = pidIdx - 1;
  var ckCol0  = ckIdx  - 1;
  var dtCol0  = dtIdx  - 1;

  var counts = [0, 0, 0, 0, 0];
  for (var r = 1; r < headerValues.length; r++) {
    var row = headerValues[r];
    if (String(row[pidCol0] || "").trim() !== patientId) continue;
    if (String(row[ckCol0]  || "").trim() !== caseKey)   continue;
    var td = row[dtCol0];
    if (!(td instanceof Date)) continue;
    var idx = (td.getFullYear() - startYear) * 12 + (td.getMonth() - startMonth);
    if (idx < 0 || idx > 4) continue;
    counts[idx]++;
  }
  return counts;
}

/**
 * caseKey 単位で金属副子等加算の通算算定回数を集計（§18.3 Phase 2）
 * 施術明細の「金属副子_確定」> 0 かつ 施術日 < beforeDate の visitKey を重複除去してカウント。
 * detailMap / detailValues が渡されなければ 0 を返す（Phase 1 モードと共存可）。
 * @param {Array[][]} detailValues - 施術明細 getDataRange().getValues()
 * @param {Object} detailMap - buildHeaderColMap_ の結果（AM_DETAIL_COLS キー）
 * @param {string} caseKey
 * @param {Date} beforeDate - この日より前の来院のみカウント（当日除外）
 * @return {number} 通算算定回数（0以上）
 */
function buildMetalCountByCaseKey_V3_(detailValues, detailMap, caseKey, beforeDate) {
  if (!detailValues || !detailMap || !caseKey) return 0;
  var ckIdx    = detailMap[AM_DETAIL_COLS.caseKey];
  var dtIdx    = detailMap[AM_DETAIL_COLS.treatDate];
  var metalIdx = detailMap[AM_DETAIL_COLS.metalOut];
  var vkIdx    = detailMap[AM_DETAIL_COLS.visitKey];
  if (!ckIdx || !dtIdx || !metalIdx || !vkIdx) return 0;

  var seen = {};  // visitKey → true（重複除去）
  for (var r = 1; r < detailValues.length; r++) {
    var row = detailValues[r];
    if (String(row[ckIdx - 1] || "").trim() !== caseKey) continue;
    var td = row[dtIdx - 1];
    if (!(td instanceof Date)) continue;
    if (beforeDate instanceof Date && td.getTime() >= beforeDate.getTime()) continue;
    var metal = Number(row[metalIdx - 1] || 0);
    if (metal <= 0) continue;
    var vk = String(row[vkIdx - 1] || "").trim();
    if (vk) seen[vk] = true;
  }
  return Object.keys(seen).length;
}

/**
 * caseKey×当月単位で柔道整復運動後療料の算定回数を集計（Phase 2）
 * 施術明細の「運動後療_確定」> 0 かつ 施術日が treatDate と同年月 かつ 施術日 < treatDate
 * の visitKey を重複除去してカウント。
 * detailMap / detailValues が渡されなければ 0 を返す（Phase 1 モードと共存可）。
 * @param {Array[][]} detailValues - 施術明細 getDataRange().getValues()
 * @param {Object} detailMap - buildHeaderColMap_ の結果（AM_DETAIL_COLS キー）
 * @param {string} caseKey
 * @param {Date} treatDate - 当日施術日（同月判定基準 かつ 当日除外の境界）
 * @return {number} 当月算定回数（0以上）
 */
function buildExerciseCountByMonth_V3_(detailValues, detailMap, caseKey, treatDate) {
  if (!detailValues || !detailMap || !caseKey) return 0;
  if (!(treatDate instanceof Date)) return 0;
  var ckIdx  = detailMap[AM_DETAIL_COLS.caseKey];
  var dtIdx  = detailMap[AM_DETAIL_COLS.treatDate];
  var exIdx  = detailMap[AM_DETAIL_COLS.exerciseOut];
  var vkIdx  = detailMap[AM_DETAIL_COLS.visitKey];
  if (!ckIdx || !dtIdx || !exIdx || !vkIdx) return 0;

  var targetYear  = treatDate.getFullYear();
  var targetMonth = treatDate.getMonth();  // 0-indexed

  var seen = {};  // visitKey → true（重複除去）
  for (var r = 1; r < detailValues.length; r++) {
    var row = detailValues[r];
    if (String(row[ckIdx - 1] || "").trim() !== caseKey) continue;
    var td = row[dtIdx - 1];
    if (!(td instanceof Date)) continue;
    // 当日除外（当日より前のみ）
    if (td.getTime() >= treatDate.getTime()) continue;
    // 同月チェック
    if (td.getFullYear() !== targetYear || td.getMonth() !== targetMonth) continue;
    var ex = Number(row[exIdx - 1] || 0);
    if (ex <= 0) continue;
    var vk = String(row[vkIdx - 1] || "").trim();
    if (vk) seen[vk] = true;
  }
  return Object.keys(seen).length;
}

/**
 * 1部位分の金額算定（object返却版）
 *
 * Ver3_core.js 側の同名関数と同じロジックだが、
 * 内訳オブジェクトを返す点が異なる。
 *
 * @param {boolean} [metalChk] - 金属副子等加算チェック（§18.3）
 * @param {number|null} [metalPriorCount] - 当日より前の通算算定回数（null=Phase 1 モード・回数制限なし）
 * @param {boolean} [exerciseChk] - 柔道整復運動後療料チェック
 * @param {number|null} [exercisePriorCount] - 当月・当日より前の算定回数（null=Phase 1 モード・回数制限なし）
 * @return {Object} { base, cold, warm, electro, taiki, metalOut, exerciseOut, coef, longTermCoef, total, ... }
 */
function calcOnePartAmount_V3_(settings, kubun, byomei, injuryDate, treatDate, coldChk, warmChk, elecChk, partOrder, reasons, bui, monthlyVisitCounts, metalChk, metalPriorCount, exerciseChk, exercisePriorCount) {
  var injuryType = detectInjuryType_V3_(byomei);
  var base = calcBaseFee_V3_(settings, kubun, injuryType, bui);

  // 骨折・不全骨折で整復料/固定料が取得できない場合
  if (base === 0 && kubun === "初検" && (injuryType === "骨折" || injuryType === "不全骨折")) {
    reasons.push("整復料/固定料 取得不可（" + (bui || "部位不明") + "：設定シートにキーがありません）");
  }

  var dayDiff = null;
  if (injuryDate instanceof Date && treatDate instanceof Date) {
    dayDiff = diffDays_V3_(injuryDate, treatDate);
  }

  var extType = detectInjuryType_V3_(byomei);

  // --- 冷罨法 §9.1 ---
  var cold = 0;
  if (coldChk) {
    var coldAllowed = false;
    if (dayDiff != null) {
      if (extType === "骨折" || extType === "不全骨折") {
        coldAllowed = (dayDiff <= 6);
      } else if (extType === "脱臼") {
        coldAllowed = (dayDiff <= 4);
      } else {
        coldAllowed = (dayDiff <= 1);
      }
    }
    if (coldAllowed) {
      cold = settings.cold;
    } else {
      reasons.push("冷罨法 算定不可（" + (byomei || "不明") + "：受傷後" + (dayDiff != null ? dayDiff : "?") + "日）");
    }
  }

  // --- 温罨法 §9.2 ---
  var warm = 0;
  if (warmChk) {
    // 初検日特例（厚労省通知 P43）: 整復/施療を行う初検日は温罨法算定不可
    if (kubun === "初検") {
      reasons.push("温罨法 算定不可（初検日特例：" + (byomei || "不明") + "）");
    } else {
      var warmAllowed = false;
      if (dayDiff != null) {
        if (extType === "骨折" || extType === "不全骨折") {
          warmAllowed = (dayDiff >= 7);
        } else {
          warmAllowed = (dayDiff >= 5);
        }
      }
      if (warmAllowed) {
        warm = settings.warm;
      } else {
        reasons.push("温罨法 算定不可（" + (byomei || "不明") + "：受傷後" + (dayDiff != null ? dayDiff : "?") + "日）");
      }
    }
  }

  // --- 電療 §9.2 ---
  var electro = 0;
  if (elecChk) {
    var elecAllowed = false;
    if (dayDiff != null) {
      if (extType === "骨折" || extType === "不全骨折") {
        elecAllowed = (dayDiff >= 7);
      } else {
        elecAllowed = (dayDiff >= 5);
      }
    }
    if (elecAllowed) {
      electro = settings.electro;
    } else {
      reasons.push("電療 算定不可（" + (byomei || "不明") + "：受傷後" + (dayDiff != null ? dayDiff : "?") + "日）");
    }
  }

  // 待機料（温/電いずれかが算定可のとき）
  var taiki = (warm > 0 || electro > 0) ? settings.taiki : 0;

  // 長期減額 §11（骨折・不全骨折は対象外）
  var ltCoef = calcLongTermCoef_V3_(injuryType, injuryDate, treatDate, monthlyVisitCounts);
  if (ltCoef === 0.50) {
    reasons.push("長期減額50%適用（" + byomei + "）");
  } else if (ltCoef < 1.0) {
    reasons.push("長期減額75%適用（" + byomei + "）");
  }
  // 継続理由書アラート: 受傷3か月超（C群・脱臼対象。骨折・不全骨折は§20対象外）
  var meAlert = calcMonthsElapsed_V3_(injuryDate, treatDate);
  if (meAlert >= 3 && injuryType !== "骨折" && injuryType !== "不全骨折") {
    reasons.push("長期施術3か月超（継続理由書確認）");
  }
  // 長期対象: 後療料(再検/後療時のbase)・冷・温・電。初検時のbaseと待機料は非対象
  var ltBase = (kubun === "初検") ? base : Math.round(base * ltCoef);
  var ltCold = Math.round(cold * ltCoef);
  var ltWarm = Math.round(warm * ltCoef);
  var ltElectro = Math.round(electro * ltCoef);

  // 多部位逓減 §10
  var coef = (partOrder >= 3) ? Number(settings.multiCoef3 || 0.6) : 1.0;

  // 金属副子等加算 §18.3（Phase 1: 傷病種別チェック / Phase 2: 回数制限追加）
  var metalOut = 0;
  if (metalChk) {
    if (injuryType === "骨折" || injuryType === "不全骨折" || injuryType === "脱臼") {
      // Phase 2: 回数制限チェック（metalPriorCount が渡された場合のみ）
      if (metalPriorCount != null && metalPriorCount >= 3) {
        reasons.push("金属副子等加算 算定上限超（通算3回）");
      } else {
        metalOut = settings.metalAddon || 0;
      }
    } else {
      // C群（打撲・捻挫・挫傷）や不明傷病は算定不可
      reasons.push("金属副子等加算 算定不可（対象外傷病：" + (byomei || "不明") + "）");
    }
  }

  // 柔道整復運動後療料（骨折・不全骨折・脱臼 / dayDiff >= 15 / 逓減対象外）
  var exerciseOut = 0;
  if (exerciseChk) {
    if (injuryType === "骨折" || injuryType === "不全骨折" || injuryType === "脱臼") {
      if (dayDiff != null && dayDiff >= 15) {
        // Phase 2: 当月5回上限チェック（exercisePriorCount が渡された場合のみ）
        if (exercisePriorCount != null && exercisePriorCount >= 5) {
          reasons.push("運動後療料 算定上限超（当月5回）");
        } else {
          exerciseOut = settings.exerciseAddon || 0;
        }
      } else {
        reasons.push("運動後療料 算定不可（受傷後15日未満）");
      }
    } else {
      // C群（打撲・捻挫・挫傷）や不明傷病は算定不可
      reasons.push("運動後療料 算定不可（対象外傷病：" + (byomei || "不明") + "）");
    }
  }

  return {
    base: base,
    cold: cold,
    warm: warm,
    electro: electro,
    taiki: taiki,
    metalOut: metalOut,       // §18.3 逓減対象外
    exerciseOut: exerciseOut, // 柔道整復運動後療料 逓減対象外
    coef: coef,
    longTermCoef: ltCoef,
    total: (ltBase + ltCold + ltWarm + ltElectro + taiki) * coef + metalOut + exerciseOut, // metal/exercise は coef・ltCoef 乗算なし
    byomei: byomei,
    partOrder: partOrder,
    injuryDate: injuryDate,
    coldChk: coldChk,
    warmChk: warmChk,
    electroChk: elecChk,
    metalChk: !!metalChk,
    exerciseChk: !!exerciseChk,
  };
}

/* =======================================================================
   calcHeaderAmountsByVisitKey_V3_  ―  金額計算（来院ケースベース）
   SPEC.md §4-1 混在来院日課金優先順位に準拠
   - 患者×月上限（初検料/相談支援料 各月1回）
   - 再検料: caseKey単位（初検後最初の後療日のみ / kubun===再検で自動判定）
   - 混在visit課金優先: hasBillableInitial > hasReexam > 0
   - 後療料は初検日以外の施術日（再検日は再検＋後療）
   - 冷温電は傷病種別×受傷日経過で算定可否判定
   - 安全弁：算定不可→金額0＋要確認TRUE＋理由記録
   ======================================================================= */
function calcHeaderAmountsByVisitKey_V3_(ss, visitKey, patientId, treatDate, kubun1, kubun2) {
  var settings = loadSettings_V3_(ss);

  var masterSh = ss.getSheetByName(SHEETS.master);
  var masterMap = buildHeaderColMap_(masterSh);
  var burden = loadBurdenRatio_V3_(masterSh, masterMap, patientId);

  var caseSh = ss.getSheetByName(SHEETS.cases);
  var caseMap = buildHeaderColMap_(caseSh);
  var headSh = ss.getSheetByName(SHEETS.header);
  var headMap = buildHeaderColMap_(headSh);

  var reasons = [];  // 要確認理由を蓄積

  // --- 患者×月 上限チェック ---
  var monthKey = fmt_(treatDate, "yyyy-MM");
  // 治癒後別負傷 [B] 判定のため caseSh / caseMap / treatDate を渡す
  var monthlyStatus = getMonthlyBilledStatus_(headSh, headMap, patientId, monthKey, visitKey, caseSh, caseMap, treatDate);

  var hasInit = (kubun1 === "初検" || kubun2 === "初検");
  var hasReexam = (kubun1 === "再検" || kubun2 === "再検");
  var hasKoryo = (kubun1 === "再検" || kubun1 === "後療" || kubun2 === "再検" || kubun2 === "後療");

  // --- 初検料 ---
  var initFee = 0;
  if (hasInit) {
    if (monthlyStatus.initBilled) {
      initFee = 0;
      reasons.push("同月別ケース初回 初検抑制");
    } else {
      initFee = settings.initFee;
    }
  }

  // 実際に算定される初検があるか（抑制された初検は含めない）
  var hasBillableInitial = (initFee > 0);

  // --- 相談支援料（初検料を算定する日のみ） ---
  var supportFee = 0;
  if (hasBillableInitial) {
    if (monthlyStatus.supportBilled) {
      supportFee = 0;
    } else {
      supportFee = settings.initSupport;
    }
  }

  // --- 再検料（caseKey単位：初検後最初の後療日のみ）---
  // kubun===再検 は calcEpisodeForCase_ が「エピソード内priorCount===1」
  // = 同一caseKeyの初検後・最初の後療日と判定した結果
  // 初検が算定される日（hasBillableInitial）は初検優先で再検料は立てない
  //
  // monthlyStatus.reBilled 制御（2026-03-18 追加）:
  //   [A] 施術継続中: 先行ケースが月内に再検算定済かつ治癒していない → reBilled=true → 当月2件目の再検を抑制
  //   [B] 治癒後別負傷: 先行ケースの再検は治癒後ケース扱い → isCaseEndedBefore_ で suppressReBilled=true
  //                   → reBilled=false → 後続エピソードの再検を許可
  var reFee = 0;
  if (hasReexam && !hasBillableInitial && !monthlyStatus.reBilled) {
    reFee = settings.reFee;
  }

  // --- 部位別明細金額 ---
  // 【設計方針】初検抑制（initFee=0）は「初検料加算を0にする」のみ。
  //   部位基本料の算定区分は case 自体の生の kubun を使う。
  //   初検抑制があっても、その日は case2 にとって初検日であり施療料を適用する。
  //   effectiveKubun は来院ヘッダ要約・chargeReason 生成専用（部位計算には使わない）。
  var calcKoryoOnThisDay = !hasBillableInitial;  // 来院ヘッダ要約生成のために保持
  var effectiveKubun1 = calcKoryoOnThisDay ? (kubun1 === "初検" ? "後療" : kubun1) : kubun1;
  var effectiveKubun2 = calcKoryoOnThisDay ? (kubun2 === "初検" ? "後療" : kubun2) : kubun2;

  var headerValuesForMvc = headSh.getDataRange().getValues();
  // §18.3 Phase 2: 施術明細を読み込み金属副子等加算の通算回数判定に使用
  var detailShForMetal = ss.getSheetByName(SHEETS.detail);
  var detailMapForMetal = buildHeaderColMap_(detailShForMetal);
  var detailValuesForMetal = detailShForMetal.getDataRange().getValues();
  // 部位基本料は生の kubun で算定（初検抑制時も初検日扱いの施療料を適用）
  var detail1 = calcCaseDetailAmount_V3_(caseSh, caseMap, visitKey, 1, kubun1, treatDate, settings, reasons, headerValuesForMvc, headMap, detailValuesForMetal, detailMapForMetal);
  var detail2 = calcCaseDetailAmount_V3_(caseSh, caseMap, visitKey, 2, kubun2, treatDate, settings, reasons, headerValuesForMvc, headMap, detailValuesForMetal, detailMapForMetal);
  var detailSum = detail1.total + detail2.total;

  // --- 来院合計 = 初検料 + 再検料 + 相談支援料 + 明細合計 ---
  var visitTotal = initFee + reFee + supportFee + detailSum;

  var unit = settings.roundUnit || 1;
  var windowPay = roundToUnit_V3_(visitTotal * burden, unit);
  var claimPay = visitTotal - windowPay;

  var needCheck = (reasons.length > 0);
  var needCheckReason = reasons.join(";");

  // ─── 新5列生成ブロック ───────────────────────────────────────────────
  // 使用する既存変数: initFee / reFee / hasBillableInitial / hasReexam /
  //                  hasKoryo / kubun1 / kubun2 / reasons

  // ── 補助フラグ ──
  var isMixed = (kubun2 != null && String(kubun2).trim() !== "");
  var initSuppressed = (reasons || []).some(function(r) {
    return r.indexOf("初検抑制") !== -1;
  });

  // ── 算定区分: 課金実績の代表区分 ──
  var billedKubun;
  if (initFee > 0) {
    billedKubun = "初検";
  } else if (reFee > 0) {
    billedKubun = "再検";
  } else if (hasKoryo) {
    billedKubun = "後療";
  } else {
    billedKubun = "算定なし";
  }

  // ── Mixed区分: 複数ケースの有無 ──
  var mixedFlag = isMixed ? "Mixed" : "通常";

  // ── case1要約 ──
  var k1 = String(kubun1 || "").trim();
  var case1Summary;
  if      (k1 === "初検") { case1Summary = "case1:初検"; }
  else if (k1 === "再検") { case1Summary = "case1:再検"; }
  else if (k1 === "後療") { case1Summary = "case1:後療"; }
  else                    { case1Summary = "case1:なし"; }

  // ── case2要約: kubun2 × 初検抑制フラグ ──
  var k2 = String(kubun2 || "").trim();
  var case2Summary;
  if (!k2) {
    case2Summary = "case2:なし";
  } else if (k2 === "初検") {
    case2Summary = initSuppressed ? "case2:初検(抑制)" : "case2:初検";
  } else if (k2 === "再検") {
    case2Summary = "case2:再検";
  } else if (k2 === "後療") {
    case2Summary = "case2:後療";
  } else {
    case2Summary = "case2:" + k2;  // 想定外区分の安全弁
  }

  // ── 課金理由要約: 優先順位7パターン ──
  var chargeReason;
  if (hasBillableInitial && !isMixed) {
    chargeReason = "初検のみ";
  } else if (hasBillableInitial && isMixed) {
    chargeReason = "算定可能な初検ありのため初検採用";       // M02
  } else if (!hasBillableInitial && reFee > 0 && isMixed && initSuppressed) {
    chargeReason = "初検抑制のため再検採用";                 // M01
  } else if (!hasBillableInitial && reFee > 0 && isMixed && !initSuppressed) {
    chargeReason = "再検ありのため再検採用";                 // 後療+再検 mixed など
  } else if (!hasBillableInitial && reFee > 0 && !isMixed) {
    chargeReason = "再検のみ";
  } else if (!hasBillableInitial && reFee === 0 && hasKoryo && isMixed && initSuppressed) {
    chargeReason = "初検抑制かつ再検対象なし";               // M03
  } else if (!hasBillableInitial && reFee === 0 && hasKoryo) {
    chargeReason = "後療のみ";                               // TC03
  } else {
    chargeReason = "算定なし";
  }
  // ──────────────────────────────────────────────────────────────────────

  return {
    initFee: initFee,
    reFee: reFee,
    supportFee: supportFee,
    detailSum: detailSum,
    visitTotal: visitTotal,
    windowPay: windowPay,
    claimPay: claimPay,
    needCheck: needCheck,
    needCheckReason: needCheckReason,
    details: { case1Parts: detail1.parts, case2Parts: detail2.parts },
    // 新5列: mixed case 説明性列
    billedKubun: billedKubun,
    mixedFlag: mixedFlag,
    case1Summary: case1Summary,
    case2Summary: case2Summary,
    chargeReason: chargeReason,
    // 抑制変換後の実効区分（upsertDetailRows_V3_ で detail の区分列に反映する）
    effectiveKubun1: effectiveKubun1,
    effectiveKubun2: effectiveKubun2,
  };
}

/** 当月の既算定状況を来院ヘッダから取得（自分自身のvisitKeyは除外） */
/**
 * 患者×月の算定済みフラグを返す。
 *
 * opt_caseSh / opt_caseMap / opt_treatDate を渡した場合、
 * 治癒後別負傷 [B] の判定を行う:
 *   月内に initFee > 0 の行があっても、そのケースが opt_treatDate より前に終了していれば
 *   initBilled=false を維持する（新規エピソードとして初検料を再算定可とする）。
 *
 * 制度根拠: 治癒後に同月内で新たな別負傷が発生した場合は初検料を再度算定できる
 *           （厚生労働省集団指導資料 / §3-6 月内上限ルール）
 */
function getMonthlyBilledStatus_(headSh, headMap, patientId, monthKey, excludeVisitKey,
                                  opt_caseSh, opt_caseMap, opt_treatDate) {
  var result = { initBilled: false, reBilled: false, supportBilled: false, initDate: null };
  var lastRow = headSh.getLastRow();
  if (lastRow < 2) return result;

  var n = lastRow - 1;
  var cPid  = headMap[HEADER_COLS.patientId];
  var cDt   = headMap[HEADER_COLS.treatDate];
  var cVk   = headMap[HEADER_COLS.visitKey];
  var cInit = headMap[HEADER_COLS.initFee];
  var cRe   = headMap[HEADER_COLS.reFee];
  var cSup  = headMap[HEADER_COLS.supportFee];
  if (!cPid || !cDt || !cVk || !cInit || !cRe || !cSup) return result;

  // 治癒後別負傷チェック用: 来院ヘッダの caseKey 列（任意）
  var cCk = headMap[HEADER_COLS.caseKey];

  var pidVals   = headSh.getRange(2, cPid,  n, 1).getValues().flat();
  var dtVals    = headSh.getRange(2, cDt,   n, 1).getValues().flat();
  var vkVals    = headSh.getRange(2, cVk,   n, 1).getValues().flat();
  var initVals  = headSh.getRange(2, cInit, n, 1).getValues().flat();
  var reVals    = headSh.getRange(2, cRe,   n, 1).getValues().flat();
  var supVals   = headSh.getRange(2, cSup,  n, 1).getValues().flat();
  var ckVals    = (cCk) ? headSh.getRange(2, cCk, n, 1).getValues().flat() : null;

  for (var i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    var d = dtVals[i];
    if (!(d instanceof Date)) continue;
    if (fmt_(d, "yyyy-MM") !== monthKey) continue;
    if (String(vkVals[i] || "").trim() === excludeVisitKey) continue;

    var iv = Number(initVals[i] || 0);
    var rv = Number(reVals[i] || 0);
    var sv = Number(supVals[i] || 0);

    if (iv > 0) {
      // 治癒後別負傷 [B] チェック:
      //   opt_caseSh / opt_treatDate が提供されており、かつこの initFee 行の caseKey が
      //   opt_treatDate より前に終了していれば → 別エピソード扱いで initBilled=true を立てない
      var suppressInitBilled = false;
      if (opt_caseSh && opt_caseMap && opt_treatDate instanceof Date && ckVals) {
        var prevCaseKey = String(ckVals[i] || "").trim();
        if (prevCaseKey) {
          suppressInitBilled = isCaseEndedBefore_(opt_caseSh, opt_caseMap, prevCaseKey, opt_treatDate);
        }
      }
      if (!suppressInitBilled) {
        result.initBilled = true;
        result.initDate = d;
      }
    }
    if (rv > 0) {
      // 治癒後別負傷 [B] チェック（initBilled と同じロジック）:
      //   この再検料行の caseKey が opt_treatDate より前に治癒済みであれば
      //   別エピソード扱いで reBilled=true を立てない（[B] 再検は月内制限対象外）
      //   case が治癒していない（施術継続中）なら [A] → reBilled=true（月内1回制限）
      var suppressReBilled = false;
      if (opt_caseSh && opt_caseMap && opt_treatDate instanceof Date && ckVals) {
        var ckForRe = String(ckVals[i] || "").trim();
        if (ckForRe) {
          suppressReBilled = isCaseEndedBefore_(opt_caseSh, opt_caseMap, ckForRe, opt_treatDate);
        }
      }
      if (!suppressReBilled) result.reBilled = true;
    }
    if (sv > 0) result.supportBilled = true;
  }
  return result;
}

/**
 * 指定 caseKey の最遅施術終了日が treatDate より前であれば true を返す（ケース治癒済み判定）。
 * 施術終了日_部位1 / 部位2 の両方を参照し、最も遅い日付を使用する。
 * 終了日が未設定の場合は false（施術継続中）。
 *
 * 制度根拠: SPEC.md §3-6 治癒後別負傷 [B] 判定の補助ロジック
 */
function isCaseEndedBefore_(caseSh, caseMap, caseKey, treatDate) {
  var v = caseSh.getDataRange().getValues();
  if (v.length < 2) return false;

  var cKey = caseMap[CASE_COLS.caseKey];
  var cE1  = caseMap[CASE_COLS.end1];
  var cE2  = caseMap[CASE_COLS.end2];
  if (cKey === undefined) return false;

  var latestEnd = null;
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][cKey - 1] || "").trim() !== caseKey) continue;  // caseMap is 1-based
    var e1 = (cE1 && v[r][cE1 - 1] instanceof Date) ? v[r][cE1 - 1] : null;
    var e2 = (cE2 && v[r][cE2 - 1] instanceof Date) ? v[r][cE2 - 1] : null;
    var eMax = (e1 && e2) ? (e1 > e2 ? e1 : e2) : (e1 || e2);
    if (eMax && (!latestEnd || eMax > latestEnd)) latestEnd = eMax;
  }
  return latestEnd instanceof Date && latestEnd < treatDate;
}

/** 当月初検の次回来院日かどうか判定 */
function checkIsNextVisitAfterMonthlyInit_(headSh, headMap, caseSh, caseMap, patientId, monthKey, treatDate) {
  // 当月の初検日を探す
  var lastRow = headSh.getLastRow();
  if (lastRow < 2) return false;

  var n = lastRow - 1;
  var cPid  = headMap[HEADER_COLS.patientId];
  var cDt   = headMap[HEADER_COLS.treatDate];
  var cInit = headMap[HEADER_COLS.initFee];
  if (!cPid || !cDt || !cInit) return false;

  var pidVals  = headSh.getRange(2, cPid, n, 1).getValues().flat();
  var dtVals   = headSh.getRange(2, cDt,  n, 1).getValues().flat();
  var initVals = headSh.getRange(2, cInit, n, 1).getValues().flat();

  var initDate = null;
  for (var i = 0; i < n; i++) {
    if (String(pidVals[i] || "").trim() !== patientId) continue;
    var d = dtVals[i];
    if (!(d instanceof Date)) continue;
    if (fmt_(d, "yyyy-MM") !== monthKey) continue;
    if (Number(initVals[i] || 0) > 0) {
      initDate = d;
      break;  // 当月最初の初検日
    }
  }
  if (!initDate) return false;

  // 当月の初検日以降の来院日を収集（初検日自体は除く）
  var visitDates = getPatientVisitDatesFromCases_(caseSh, caseMap, patientId);
  var afterInitDates = visitDates.filter(function(d) {
    return d.getTime() > initDate.getTime() && fmt_(d, "yyyy-MM") === monthKey;
  }).sort(function(a, b) { return a.getTime() - b.getTime(); });

  if (!afterInitDates.length) return false;

  // 次回来院日 = 初検日の直後の来院日
  var nextVisitDate = afterInitDates[0];
  return sameDateKey_(nextVisitDate) === sameDateKey_(treatDate);
}

/** 1ケース分の明細金額を来院ケースの部位データから算定（SPEC準拠版）
 *  @param {Array[][]|null} [detailValues] - 施術明細全データ（Phase 2 回数制限用。null=Phase 1 モード）
 *  @param {Object|null} [detailMap] - 施術明細列マップ（Phase 2 回数制限用。null=Phase 1 モード）
 *  @return {{ total: number, parts: Object[] }} 合計と部位別内訳配列
 */
function calcCaseDetailAmount_V3_(caseSh, caseMap, visitKey, caseNo, kubun, treatDate, settings, reasons, headerValues, headMap, detailValues, detailMap) {
  // 来院ケース行は visitKey + caseNo で検索
  var rowIndex = findCaseRowByVisitKeyAndCaseNo_(caseSh, caseMap, visitKey, caseNo);
  if (rowIndex === 0) return { total: 0, parts: [] };

  var row = caseSh.getRange(rowIndex, 1, 1, caseSh.getLastColumn()).getValues()[0];
  var get = function(name) { return row[caseMap[name] - 1]; };

  var patientIdVal = String(get(CASE_COLS.patientId) || "").trim();
  var caseKeyVal   = String(get(CASE_COLS.caseKey)   || "").trim();

  var total = 0;
  var partCount = 0;
  var parts = [];

  // 部位1 ★終了日を超えた来院ではスキップ
  var p1 = String(get(CASE_COLS.p1) || "").trim();
  var d1 = String(get(CASE_COLS.d1) || "").trim();
  var inj1 = get(CASE_COLS.inj1);
  var end1 = get(CASE_COLS.end1);
  var isEnded1 = (end1 instanceof Date) && (treatDate instanceof Date)
    && (end1.getTime() < treatDate.getTime());
  if ((p1 || d1 || (inj1 instanceof Date)) && !isEnded1) {
    partCount++;
    var mvc1 = buildMonthlyVisitCounts_V3_(headerValues, headMap, patientIdVal, caseKeyVal, inj1 instanceof Date ? inj1 : null);
    var metal1Chk = get(CASE_COLS.metal1) === true;
    var metalCount1 = (metal1Chk && detailValues && detailMap)
      ? buildMetalCountByCaseKey_V3_(detailValues, detailMap, caseKeyVal, treatDate)
      : null;  // null = Phase 1 モード（回数制限なし）
    var exercise1Chk = get(CASE_COLS.exercise1) === true;
    var exerciseCount1 = (exercise1Chk && detailValues && detailMap)
      ? buildExerciseCountByMonth_V3_(detailValues, detailMap, caseKeyVal, treatDate)
      : null;  // null = Phase 1 モード（回数制限なし）
    var part1 = calcOnePartAmount_V3_(settings, kubun, d1, inj1, treatDate,
      get(CASE_COLS.cold1) === true,
      get(CASE_COLS.warm1) === true,
      get(CASE_COLS.elec1) === true,
      partCount, reasons, p1, mvc1,
      metal1Chk, metalCount1, exercise1Chk, exerciseCount1);  // §18.3 + 運動後療料 Phase 1+2
    part1.bui = p1;
    total += part1.total;
    parts.push(part1);
  }

  // 部位2 ★終了日を超えた来院ではスキップ
  var p2 = String(get(CASE_COLS.p2) || "").trim();
  var d2 = String(get(CASE_COLS.d2) || "").trim();
  var inj2 = get(CASE_COLS.inj2);
  var end2 = get(CASE_COLS.end2);
  var isEnded2 = (end2 instanceof Date) && (treatDate instanceof Date)
    && (end2.getTime() < treatDate.getTime());
  if ((p2 || d2 || (inj2 instanceof Date)) && !isEnded2) {
    partCount++;
    var mvc2 = buildMonthlyVisitCounts_V3_(headerValues, headMap, patientIdVal, caseKeyVal, inj2 instanceof Date ? inj2 : null);
    var metal2Chk = get(CASE_COLS.metal2) === true;
    var metalCount2 = (metal2Chk && detailValues && detailMap)
      ? buildMetalCountByCaseKey_V3_(detailValues, detailMap, caseKeyVal, treatDate)
      : null;
    var exercise2Chk = get(CASE_COLS.exercise2) === true;
    var exerciseCount2 = (exercise2Chk && detailValues && detailMap)
      ? buildExerciseCountByMonth_V3_(detailValues, detailMap, caseKeyVal, treatDate)
      : null;
    var part2 = calcOnePartAmount_V3_(settings, kubun, d2, inj2, treatDate,
      get(CASE_COLS.cold2) === true,
      get(CASE_COLS.warm2) === true,
      get(CASE_COLS.elec2) === true,
      partCount, reasons, p2, mvc2,
      metal2Chk, metalCount2, exercise2Chk, exerciseCount2);  // §18.3 + 運動後療料 Phase 1+2
    part2.bui = p2;
    total += part2.total;
    parts.push(part2);
  }

  return { total: total, parts: parts };
}

/**
 * visitKey単位で、施術明細→来院ヘッダを再計算（SPEC.md準拠版）
 *
 * 冷罨法ルール（§9.1）:
 *   打撲/捻挫: dayDiff≤1
 *   脱臼: dayDiff≤4
 *   骨折/不全骨折: dayDiff≤6
 *
 * 温罨法/電療ルール（§9.2）:
 *   打撲/捻挫/脱臼: dayDiff≥5
 *   骨折/不全骨折: dayDiff≥7
 *
 * 後療料: 初検日以外の施術日（再検/後療）に算定
 * 安全弁: 算定不可→金額0、チェック残す、要確認TRUE、理由記録
 */
function recalcAmountsByVisitKey_V3_(ss, visitKey) {
  var settings = loadSettings_V3_(ss);

  var detailSh = ss.getSheetByName(SHEETS.detail);
  var headerSh = ss.getSheetByName(SHEETS.header);
  var masterSh = ss.getSheetByName(SHEETS.master);

  var maps = {
    detail: buildHeaderColMap_(detailSh),
    header: buildHeaderColMap_(headerSh),
    master: buildHeaderColMap_(masterSh),
  };

  // 必須列チェック（不足なら即エラー＝事故ゼロ）
  ensureRequiredCols_(maps.detail, Object.values(AM_DETAIL_COLS), SHEETS.detail);
  ensureRequiredCols_(maps.header, [
    HEADER_COLS.visitKey,
    HEADER_COLS.initFee,
    HEADER_COLS.reFee,
    HEADER_COLS.supportFee,
    HEADER_COLS.detailSum,
    HEADER_COLS.visitTotal,
    HEADER_COLS.windowPay,
    HEADER_COLS.claimPay,
  ], SHEETS.header);
  ensureRequiredCols_(maps.master, [MASTER_COLS.patientId, MASTER_COLS.burden], SHEETS.master);

  // 明細全取得
  var detailValues = detailSh.getDataRange().getValues();
  if (detailValues.length < 2) throw new Error("施術明細にデータがありません。");
  // 来院ヘッダ全取得（月別来院回数集計用 §11）
  var headerValuesAll = headerSh.getDataRange().getValues();

  // visitKey該当行を収集（0-based index）
  var vkCol0 = maps.detail[AM_DETAIL_COLS.visitKey] - 1;
  var rows0 = [];
  for (var r0 = 1; r0 < detailValues.length; r0++) {
    if (String(detailValues[r0][vkCol0] || "").trim() === visitKey) rows0.push(r0);
  }
  if (!rows0.length) throw new Error("施術明細で visitKey=" + visitKey + " が見つかりません。");

  // 患者ID
  var pidCol0 = maps.detail[AM_DETAIL_COLS.patientId] - 1;
  var patientId = String(detailValues[rows0[0]][pidCol0] || "").trim();
  if (!patientId) throw new Error("施術明細の患者IDが空です。");

  // 負担割合
  var burden = loadBurdenRatio_V3_(masterSh, maps.master, patientId);

  var total = 0;
  var reasons = [];

  // 明細行ごとに金額算定＆書き込み
  for (var i = 0; i < rows0.length; i++) {
    var r0 = rows0[i];
    var row = detailValues[r0];

    var kubun = String(row[maps.detail[AM_DETAIL_COLS.kubun] - 1] || "").trim();
    var byomei = String(row[maps.detail[AM_DETAIL_COLS.byomei] - 1] || "").trim();

    var treatDate = asDate_V3_(row[maps.detail[AM_DETAIL_COLS.treatDate] - 1]);
    var injuryDate = pickInjuryDate_V3_(row, maps.detail);

    var partOrder = Number(row[maps.detail[AM_DETAIL_COLS.partOrder] - 1] || 0) || 0;

    var coldChk = row[maps.detail[AM_DETAIL_COLS.coldChk] - 1] === true;
    var warmChk = row[maps.detail[AM_DETAIL_COLS.warmChk] - 1] === true;
    var electroChk = row[maps.detail[AM_DETAIL_COLS.electroChk] - 1] === true;
    var metalChkVal = row[maps.detail[AM_DETAIL_COLS.metalChk] - 1] === true;  // §18.3
    var exerciseChkVal = maps.detail[AM_DETAIL_COLS.exerciseChk]
      ? row[maps.detail[AM_DETAIL_COLS.exerciseChk] - 1] === true
      : false;

    var buiVal = String(row[maps.detail[AM_DETAIL_COLS.bui] - 1] || "").trim();
    var caseKeyVal = String(row[maps.detail[AM_DETAIL_COLS.caseKey] - 1] || "").trim();
    var mvc = buildMonthlyVisitCounts_V3_(headerValuesAll, maps.header, patientId, caseKeyVal, injuryDate);
    // §18.3 Phase 2: 通算回数（caseKey×当日より前）
    var metalPriorCount = metalChkVal
      ? buildMetalCountByCaseKey_V3_(detailValues, maps.detail, caseKeyVal, treatDate)
      : null;
    // 運動後療料 Phase 2: 当月回数（caseKey×当月×当日より前）
    var exercisePriorCount = exerciseChkVal
      ? buildExerciseCountByMonth_V3_(detailValues, maps.detail, caseKeyVal, treatDate)
      : null;
    var part = calcOnePartAmount_V3_(settings, kubun, byomei, injuryDate, treatDate,
      coldChk, warmChk, electroChk, partOrder, reasons, buiVal, mvc, metalChkVal, metalPriorCount, exerciseChkVal, exercisePriorCount);

    // 相談支援：運用ON列が無いので事故防止で0固定
    var support = 0;

    total += part.total;

    // 書き込み（1-based行/列）
    var row1 = r0 + 1;
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.coefOut]).setValue(part.coef);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.baseOut]).setValue(part.base);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.supportOut]).setValue(support);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.coldOut]).setValue(part.cold);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.warmOut]).setValue(part.warm);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.electroOut]).setValue(part.electro);
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.taikiOut]).setValue(part.taiki);
    if (maps.detail[AM_DETAIL_COLS.metalOut]) {
      detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.metalOut]).setValue(part.metalOut);
    }
    if (maps.detail[AM_DETAIL_COLS.exerciseOut]) {
      detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.exerciseOut]).setValue(part.exerciseOut);
    }
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.rowTotalOut]).setValue(part.total);
  }

  // ヘッダ行を探す（detailSum更新前に参照してinitFee/reFee/supportFeeを取得）
  // ★ headerValuesAll は月別来院回数集計用に上部でロード済み（再取得不要）
  var hkCol0 = maps.header[HEADER_COLS.visitKey] - 1;
  var headerRow0 = -1;
  for (var r0 = 1; r0 < headerValuesAll.length; r0++) {
    if (String(headerValuesAll[r0][hkCol0] || "").trim() === visitKey) {
      headerRow0 = r0;
      break;
    }
  }
  if (headerRow0 === -1) throw new Error("来院ヘッダで visitKey=" + visitKey + " が見つかりません。");

  // HIGH-1修正: initFee/reFee/supportFee はヘッダの既存値を保持し、
  // detailSum のみ再計算した上で visitTotal を再構成する。
  // SPEC §13: visitTotal = initFee + reFee + supportFee + detailSum
  var headerRow = headerValuesAll[headerRow0];
  var existingInitFee   = Number(headerRow[maps.header[HEADER_COLS.initFee]   - 1] || 0);
  var existingReFee     = Number(headerRow[maps.header[HEADER_COLS.reFee]     - 1] || 0);
  var existingSupportFee = Number(headerRow[maps.header[HEADER_COLS.supportFee] - 1] || 0);
  var detailSum = total;  // total = 明細合計（部位計）
  var visitTotal = existingInitFee + existingReFee + existingSupportFee + detailSum;

  // 窓口・請求
  var unit = settings.roundUnit || 1;
  var copay = roundToUnit_V3_(visitTotal * burden, unit);
  var claim = visitTotal - copay;

  var hr1 = headerRow0 + 1;
  headerSh.getRange(hr1, maps.header[HEADER_COLS.detailSum]).setValue(detailSum);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.visitTotal]).setValue(visitTotal);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.windowPay]).setValue(copay);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.claimPay]).setValue(claim);

  // 要確認フラグ・理由の更新
  if (maps.header[HEADER_COLS.needCheck]) {
    headerSh.getRange(hr1, maps.header[HEADER_COLS.needCheck]).setValue(reasons.length > 0);
  }
  if (maps.header[HEADER_COLS.needCheckReason]) {
    headerSh.getRange(hr1, maps.header[HEADER_COLS.needCheckReason]).setValue(reasons.join(";"));
  }

  return { updatedRows: rows0.length, total: visitTotal, copay: copay, claim: claim };
}
