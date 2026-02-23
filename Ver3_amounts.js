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
 * 長期減額係数の算定（§11）
 * 骨折・不全骨折は対象外。受傷日の起算月から5か月超で75%。
 * @return {number} 1.0（減額なし）or 0.75（長期75%）
 */
function calcLongTermCoef_V3_(injuryType, injuryDate, treatDate) {
  // 骨折・不全骨折は長期減額の対象外
  if (injuryType === "骨折" || injuryType === "不全骨折") return 1.0;
  // 受傷日・施術日が不明なら減額なし
  if (!(injuryDate instanceof Date) || !(treatDate instanceof Date)) return 1.0;

  // 起算月を計算（受傷日が16日以降なら翌月起算）
  var startYear = injuryDate.getFullYear();
  var startMonth = injuryDate.getMonth(); // 0-indexed
  if (injuryDate.getDate() >= 16) {
    startMonth++;
    if (startMonth > 11) { startMonth = 0; startYear++; }
  }

  // 施術日の年月
  var treatYear = treatDate.getFullYear();
  var treatMonth = treatDate.getMonth();

  // 起算月からの経過月数
  var monthsElapsed = (treatYear - startYear) * 12 + (treatMonth - startMonth);

  // 5か月超（6か月目以降）→ 75%
  if (monthsElapsed >= 5) return 0.75;
  return 1.0;
}

/**
 * 1部位分の金額算定（object返却版）
 *
 * Ver3_core.js 側の同名関数と同じロジックだが、
 * 内訳オブジェクトを返す点が異なる。
 *
 * @return {Object} { base, cold, warm, electro, taiki, coef, longTermCoef, total, byomei, partOrder, coldChk, warmChk, electroChk, injuryDate }
 */
function calcOnePartAmount_V3_(settings, kubun, byomei, injuryDate, treatDate, coldChk, warmChk, elecChk, partOrder, reasons, bui) {
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
  var ltCoef = calcLongTermCoef_V3_(injuryType, injuryDate, treatDate);
  if (ltCoef < 1.0) {
    reasons.push("長期減額75%適用（" + byomei + "）");
  }
  // 長期対象: 後療料(再検/後療時のbase)・冷・温・電。初検時のbaseと待機料は非対象
  var ltBase = (kubun === "初検") ? base : Math.round(base * ltCoef);
  var ltCold = Math.round(cold * ltCoef);
  var ltWarm = Math.round(warm * ltCoef);
  var ltElectro = Math.round(electro * ltCoef);

  // 多部位逓減 §10
  var coef = (partOrder >= 3) ? Number(settings.multiCoef3 || 0.6) : 1.0;

  return {
    base: base,
    cold: cold,
    warm: warm,
    electro: electro,
    taiki: taiki,
    coef: coef,
    longTermCoef: ltCoef,
    total: (ltBase + ltCold + ltWarm + ltElectro + taiki) * coef,
    byomei: byomei,
    partOrder: partOrder,
    injuryDate: injuryDate,
    coldChk: coldChk,
    warmChk: warmChk,
    electroChk: elecChk,
  };
}

/* =======================================================================
   calcHeaderAmountsByVisitKey_V3_  ―  金額計算（来院ケースベース）
   SPEC.md 完全準拠版
   - 患者×月上限（初検料/再検料/相談支援料 各月1回）
   - 30日ルール（区分はcalcEpisodeForCase_で確定済み）
   - 再検料は当月初検の次回来院日のみ
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
  var monthlyStatus = getMonthlyBilledStatus_(headSh, headMap, patientId, monthKey, visitKey);

  var hasInit = (kubun1 === "初検" || kubun2 === "初検");
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

  // --- 相談支援料（初検料を算定する日のみ） ---
  var supportFee = 0;
  if (initFee > 0) {
    if (monthlyStatus.supportBilled) {
      supportFee = 0;
    } else {
      supportFee = settings.initSupport;
    }
  }

  // --- 再検料（当月初検の次回来院日のみ＋月1回上限） ---
  var reFee = 0;
  if (hasKoryo && !hasInit) {
    if (monthlyStatus.reBilled) {
      reFee = 0;
    } else {
      var isNextVisitAfterInit = checkIsNextVisitAfterMonthlyInit_(
        headSh, headMap, caseSh, caseMap, patientId, monthKey, treatDate
      );
      if (isNextVisitAfterInit) {
        reFee = settings.reFee;
      }
    }
  }

  // --- 後療料（初検日以外に算定、再検日も算定） ---
  var isInitDay = (initFee > 0);
  var isSuppressedInit = (hasInit && monthlyStatus.initBilled);

  // --- 部位別明細金額（後療料＋冷温電） ---
  var calcKoryoOnThisDay = !isInitDay || isSuppressedInit;
  var effectiveKubun1 = calcKoryoOnThisDay ? (kubun1 === "初検" ? "後療" : kubun1) : kubun1;
  var effectiveKubun2 = calcKoryoOnThisDay ? (kubun2 === "初検" ? "後療" : kubun2) : kubun2;

  var detail1 = calcCaseDetailAmount_V3_(caseSh, caseMap, visitKey, 1, effectiveKubun1, treatDate, settings, reasons);
  var detail2 = calcCaseDetailAmount_V3_(caseSh, caseMap, visitKey, 2, effectiveKubun2, treatDate, settings, reasons);
  var detailSum = detail1.total + detail2.total;

  // --- 来院合計 = 初検料 + 再検料 + 相談支援料 + 明細合計 ---
  var visitTotal = initFee + reFee + supportFee + detailSum;

  var unit = settings.roundUnit || 1;
  var windowPay = roundToUnit_V3_(visitTotal * burden, unit);
  var claimPay = visitTotal - windowPay;

  var needCheck = (reasons.length > 0);
  var needCheckReason = reasons.join(";");

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
    details: { case1Parts: detail1.parts, case2Parts: detail2.parts }
  };
}

/** 当月の既算定状況を来院ヘッダから取得（自分自身のvisitKeyは除外） */
function getMonthlyBilledStatus_(headSh, headMap, patientId, monthKey, excludeVisitKey) {
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

  var pidVals  = headSh.getRange(2, cPid, n, 1).getValues().flat();
  var dtVals   = headSh.getRange(2, cDt,  n, 1).getValues().flat();
  var vkVals   = headSh.getRange(2, cVk,  n, 1).getValues().flat();
  var initVals = headSh.getRange(2, cInit, n, 1).getValues().flat();
  var reVals   = headSh.getRange(2, cRe,   n, 1).getValues().flat();
  var supVals  = headSh.getRange(2, cSup,  n, 1).getValues().flat();

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
      result.initBilled = true;
      result.initDate = d;
    }
    if (rv > 0) result.reBilled = true;
    if (sv > 0) result.supportBilled = true;
  }
  return result;
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
 *  @return {{ total: number, parts: Object[] }} 合計と部位別内訳配列
 */
function calcCaseDetailAmount_V3_(caseSh, caseMap, visitKey, caseNo, kubun, treatDate, settings, reasons) {
  // 来院ケース行は visitKey + caseNo で検索
  var rowIndex = findCaseRowByVisitKeyAndCaseNo_(caseSh, caseMap, visitKey, caseNo);
  if (rowIndex === 0) return { total: 0, parts: [] };

  var row = caseSh.getRange(rowIndex, 1, 1, caseSh.getLastColumn()).getValues()[0];
  var get = function(name) { return row[caseMap[name] - 1]; };

  var total = 0;
  var partCount = 0;
  var parts = [];

  // 部位1
  var p1 = String(get(CASE_COLS.p1) || "").trim();
  var d1 = String(get(CASE_COLS.d1) || "").trim();
  var inj1 = get(CASE_COLS.inj1);
  if (p1 || d1 || (inj1 instanceof Date)) {
    partCount++;
    var part1 = calcOnePartAmount_V3_(settings, kubun, d1, inj1, treatDate,
      get(CASE_COLS.cold1) === true,
      get(CASE_COLS.warm1) === true,
      get(CASE_COLS.elec1) === true,
      partCount, reasons, p1);
    part1.bui = p1;
    total += part1.total;
    parts.push(part1);
  }

  // 部位2
  var p2 = String(get(CASE_COLS.p2) || "").trim();
  var d2 = String(get(CASE_COLS.d2) || "").trim();
  var inj2 = get(CASE_COLS.inj2);
  if (p2 || d2 || (inj2 instanceof Date)) {
    partCount++;
    var part2 = calcOnePartAmount_V3_(settings, kubun, d2, inj2, treatDate,
      get(CASE_COLS.cold2) === true,
      get(CASE_COLS.warm2) === true,
      get(CASE_COLS.elec2) === true,
      partCount, reasons, p2);
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
    HEADER_COLS.visitTotal,
    HEADER_COLS.windowPay,
    HEADER_COLS.claimPay,
  ], SHEETS.header);
  ensureRequiredCols_(maps.master, [MASTER_COLS.patientId, MASTER_COLS.burden], SHEETS.master);

  // 明細全取得
  var detailValues = detailSh.getDataRange().getValues();
  if (detailValues.length < 2) throw new Error("施術明細にデータがありません。");

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

    var buiVal = String(row[maps.detail[AM_DETAIL_COLS.bui] - 1] || "").trim();
    var part = calcOnePartAmount_V3_(settings, kubun, byomei, injuryDate, treatDate,
      coldChk, warmChk, electroChk, partOrder, reasons, buiVal);

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
    detailSh.getRange(row1, maps.detail[AM_DETAIL_COLS.rowTotalOut]).setValue(part.total);
  }

  // 窓口・請求
  var unit = settings.roundUnit || 1;
  var copayRaw = total * burden;
  var copay = roundToUnit_V3_(copayRaw, unit);
  var claim = total - copay;

  // ヘッダ行を探して更新
  var headerValues = headerSh.getDataRange().getValues();
  var hkCol0 = maps.header[HEADER_COLS.visitKey] - 1;
  var headerRow0 = -1;
  for (var r0 = 1; r0 < headerValues.length; r0++) {
    if (String(headerValues[r0][hkCol0] || "").trim() === visitKey) {
      headerRow0 = r0;
      break;
    }
  }
  if (headerRow0 === -1) throw new Error("来院ヘッダで visitKey=" + visitKey + " が見つかりません。");

  var hr1 = headerRow0 + 1;
  headerSh.getRange(hr1, maps.header[HEADER_COLS.visitTotal]).setValue(total);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.windowPay]).setValue(copay);
  headerSh.getRange(hr1, maps.header[HEADER_COLS.claimPay]).setValue(claim);

  // 要確認フラグ・理由の更新
  if (maps.header[HEADER_COLS.needCheck]) {
    headerSh.getRange(hr1, maps.header[HEADER_COLS.needCheck]).setValue(reasons.length > 0);
  }
  if (maps.header[HEADER_COLS.needCheckReason]) {
    headerSh.getRange(hr1, maps.header[HEADER_COLS.needCheckReason]).setValue(reasons.join(";"));
  }

  return { updatedRows: rows0.length, total: total, copay: copay, claim: claim };
}
