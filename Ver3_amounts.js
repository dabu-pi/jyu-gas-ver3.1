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
  multiCoef3: "多部位_3部位目係数",
  roundUnit: "窓口端数単位",
};

/** ===== 施術明細：列名（最終ヘッダー前提）===== */
const AM_DETAIL_COLS = {
  // key
  visitKey: "visitKey",
  patientId: "患者ID",
  treatDate: "施術日",
  kubun: "区分",
  injuryDateFixed: "受傷日_確定",
  injuryDateInput: "受傷日(入力)",
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

/** 基本料（設定シート通り） */
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

/** 傷病名→打撲/捻挫/挫傷 判別 */
function detectInjuryType_V3_(byomei) {
  if (!byomei) return null;
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
    var coef = (partOrder >= 3) ? Number(settings.multiCoef3 || 0.6) : 1.0;

    var coldChk = row[maps.detail[AM_DETAIL_COLS.coldChk] - 1] === true;
    var warmChk = row[maps.detail[AM_DETAIL_COLS.warmChk] - 1] === true;
    var electroChk = row[maps.detail[AM_DETAIL_COLS.electroChk] - 1] === true;

    var injuryType = detectInjuryType_V3_(byomei);

    // 後療料: 後療/再検の日に算定（初検日は基本料を施療料として別途算定済み）
    var base = calcBaseFee_V3_(settings, kubun, injuryType);

    // 相談支援：運用ON列が無いので事故防止で0固定
    var support = 0;

    var dayDiff = diffDays_V3_(injuryDate, treatDate);

    // 拡張傷病種別判定（脱臼・骨折対応）
    var extType = detectExtendedInjuryType_V3_(byomei);

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
    if (electroChk) {
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

    var rowTotal = (base + support + cold + warm + electro + taiki) * coef;
    total += rowTotal;

    // 書き込み（1-based行/列）
    var row1 = r0 + 1;
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
