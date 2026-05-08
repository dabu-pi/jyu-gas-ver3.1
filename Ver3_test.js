/****************************************************
 * Ver3_test.js — JREC-01 fixture テストランナー
 *
 * ★使い方
 *   Apps Script エディタで以下を実行する:
 *   - runFixtureSuite_()      全 fixture を一括実行（結果をアラートで表示）
 *   - runFixtureTest_("M01")  個別実行
 *
 * ★設計方針
 *   - SpreadsheetApp を使わない純粋計算テスト
 *   - calcOnePartAmount_V3_（Ver3_amounts.js）を直接呼び出す
 *   - computeAmountsFromFixture_V3_ で fixture → amounts 変換を担う
 *   - production コードの変更時はこちらの同等ロジックも更新すること
 *
 * ★対象: tests/jrec01/fixtures/ + expected/ の JSON を JS オブジェクトとして埋め込み
 ****************************************************/


/* =======================================================================
   テスト設定単価（設定シートの値と同値に保つこと）
   ======================================================================= */
var TEST_SETTINGS_ = {
  initFee:              1550,
  initSupport:          100,
  reFee:                410,
  shoryoDaboku:         760,
  shoryoNenZa:          760,
  shoryoZasyo:          760,
  koryoDaboku:          505,
  koryoNenZa:           505,
  koryoZasyo:           505,
  seifukuDakkyu:        5200, // 2026-03-17 設定シート確認済み
  koryoDakkyu:          720,  // 2026-03-17 設定シート確認済み
  koryoKossetu:         850,
  koryoFuzenKossetu:    720,
  cold:                 85,
  warm:                 75,   // 2026-03-17 設定シート確認済み
  electro:              33,   // 2026-03-17 設定シート確認済み
  taiki:                5,    // 2026-03-17 設定シート確認済み
  multiCoef3:           0.6,
  roundUnit:            10,
  metalAddon:           1000, // §18.3 骨折・不全骨折・脱臼 1,000円
  exerciseAddon:        320,  // 柔道整復運動後療料 骨折・不全骨折・脱臼 dayDiff>=15
  _rawMap: {
    // 整復料（骨折 初検・部位別）§17 付録A
    "整復料_骨折_鎖骨":  5500,
    "整復料_骨折_肋骨":  5500,
    "整復料_骨折_指_趾": 5500,
    "整復料_骨折_前腕":  7200,
    "整復料_骨折_上腕":  7800,
    "整復料_骨折_下腿":  7800,
    "整復料_骨折_大腿":  11800,
    // 固定料（不全骨折 初検・部位別）§17 付録A
    "固定料_鎖骨":  3000,
    "固定料_肋骨":  3000,
    "固定料_指_趾": 3000,
    "固定料_前腕":  4100,
    "固定料_上腕":  4600,
    "固定料_下腿":  4600,
    "固定料_大腿":  7200,
  },
};


/* =======================================================================
   Fixture データ（tests/jrec01/fixtures/*.json と同内容）
   ======================================================================= */
var JREC01_FIXTURES_ = {

  "TC01": {
    testId: "TC01",
    context: { patientId: "P001", treatDate: "2026-02-03",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-03", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  "TC02": {
    testId: "TC02",
    context: { patientId: "P001", treatDate: "2026-02-05",
      monthlyStatus: { initBilled: true, reBilled: false, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "再検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-03", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "TC03": {
    testId: "TC03",
    context: { patientId: "P001", treatDate: "2026-02-07",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-03", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "M01": {
    testId: "M01",
    context: { patientId: "P001", treatDate: "2026-02-10",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-03", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-10", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "M02": {
    testId: "M02",
    context: { patientId: "P001", treatDate: "2026-02-03",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "再検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-01-20", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-03", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "M03": {
    testId: "M03",
    context: { patientId: "P001", treatDate: "2026-02-10",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-10", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "M04": {
    testId: "M04",
    context: { patientId: "P001", treatDate: "2026-02-10",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-10", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-10", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "M05": {
    testId: "M05",
    context: { patientId: "P001", treatDate: "2026-02-12",
      monthlyStatus: { initBilled: true, reBilled: false, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "再検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-08", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── M06b: 治癒後別負傷（case2初検、case1は治癒済）──────────────────────────
  // treatDate=2026-02-15: case1は2/01初検・2/04再検・2/10治癒。case2は2/15新規初検（治癒後別負傷）。
  // initBilled=false: getMonthlyBilledStatus_+isCaseEndedBefore_ が確定（case1終了2/10 < treatDate2/15）
  // reBilled=true: case1の再検(2/04)が当月算定済。この来院は hasReexam=false（case2=初検）なので
  //   reBilled は reFee に影響しない。[A] 施術継続中の再検抑制は TC09b を参照。
  "M06b": {
    testId: "M06b",
    context: { patientId: "P001", treatDate: "2026-02-15",
      monthlyStatus: { initBilled: false, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-15", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC04: 30日境界 ─────────────────────────────────────────────────────
  "TC04a": {
    testId: "TC04a",
    context: { patientId: "P001", treatDate: "2026-02-04",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-01-05", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "TC04b": {
    testId: "TC04b",
    context: { patientId: "P001", treatDate: "2026-02-05",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-05", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC05: 冷罨法 打撲/捻挫 0-1日のみ ────────────────────────────────────
  "TC05a": {
    testId: "TC05a",
    context: { patientId: "P001", treatDate: "2026-02-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  "TC05b": {
    testId: "TC05b",
    context: { patientId: "P001", treatDate: "2026-02-02",
      monthlyStatus: { initBilled: true, reBilled: false, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "再検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  "TC05c": {
    testId: "TC05c",
    context: { patientId: "P001", treatDate: "2026-02-03",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC06: 温/電 捻挫 5日以降 ─────────────────────────────────────────────
  "TC06a": {
    testId: "TC06a",
    context: { patientId: "P001", treatDate: "2026-02-05",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: true, electro: true }
      ]}
    ]
  },

  "TC06b": {
    testId: "TC06b",
    context: { patientId: "P001", treatDate: "2026-02-06",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: true, electro: true }
      ]}
    ]
  },

  // ── TC07: 温/電 骨折 7日以降 ─────────────────────────────────────────────
  "TC07a": {
    testId: "TC07a",
    context: { patientId: "P001", treatDate: "2026-02-07",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "前腕", byomei: "骨折", injuryDate: "2026-02-01", cold: false, warm: true, electro: true }
      ]}
    ]
  },

  "TC07b": {
    testId: "TC07b",
    context: { patientId: "P001", treatDate: "2026-02-08",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "前腕", byomei: "骨折", injuryDate: "2026-02-01", cold: false, warm: true, electro: true }
      ]}
    ]
  },

  // ── TC08: 冷罨法 脱臼 0-4日のみ ─────────────────────────────────────────
  "TC08a": {
    testId: "TC08a",
    context: { patientId: "P001", treatDate: "2026-02-05",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "肩関節", byomei: "脱臼", injuryDate: "2026-02-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  "TC08b": {
    testId: "TC08b",
    context: { patientId: "P001", treatDate: "2026-02-06",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "肩関節", byomei: "脱臼", injuryDate: "2026-02-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC09: 月内再検消化後（両ケース後療） ────────────────────────────────
  "TC09": {
    testId: "TC09",
    context: { patientId: "P001", treatDate: "2026-02-12",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "後療", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-10", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC09b: [A]施術継続中・case2再検抑制（reBilled=true → reFee=0） ─────────
  // case1後療（施術継続中）+ case2再検（月内2件目 → reBilled=true で抑制）
  // 2026-03-18 修正: amounts.js reFee path に !reBilled 追加で正しく抑制されることを確認
  "TC09b": {
    testId: "TC09b",
    context: { patientId: "P001", treatDate: "2026-02-12",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "再検", parts: [
        { bui: "肩関節", byomei: "打撲", injuryDate: "2026-02-08", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC10: 複合（初検抑制 + 冷不可 + 温電可） ────────────────────────────
  "TC10": {
    testId: "TC10",
    context: { patientId: "P001", treatDate: "2026-02-10",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]},
      { caseNo: 2, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "捻挫", injuryDate: "2026-02-01", cold: true, warm: true, electro: true }
      ]}
    ]
  },

  // ── TC11: 初検 脱臼（整復料算定） ────────────────────────────────────────
  "TC11": {
    testId: "TC11",
    context: { patientId: "P001", treatDate: "2026-02-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "肩関節", byomei: "脱臼", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC12: 多部位逓減 2部位（係数1.0×2） ─────────────────────────────────
  "TC12": {
    testId: "TC12",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部",   byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false },
        { bui: "肩関節", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC13: 多部位逓減 3部位（1,2部位目×1.0、3部位目×0.6） ────────────────
  "TC13": {
    testId: "TC13",
    context: { patientId: "P001", treatDate: "2026-03-02",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部",   byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false },
        { bui: "肩関節", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false },
        { bui: "前腕",   byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC14: 長期逓減境界（月数ベース・受傷日1日起算） ──────────────────────
  // injuryDate=2026-02-01（16日未満→2月起算）
  // TC14a: treatDate=2026-06-30 → monthsElapsed=4 → ltCoef=1.0
  // TC14b: treatDate=2026-07-01 → monthsElapsed=5 → ltCoef=0.75
  "TC14a": {
    testId: "TC14a",
    context: { patientId: "P001", treatDate: "2026-06-30",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "TC14b": {
    testId: "TC14b",
    context: { patientId: "P001", treatDate: "2026-07-01",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC15: 不全骨折冷罨法 dayDiff境界（骨折/不全骨折: dayDiff<=6 OK, >=7 NG） ──
  // injuryDate=2026-03-01
  // TC15a: treatDate=2026-03-07 → dayDiff=6 → coldAllowed=true → cold=85
  // TC15b: treatDate=2026-03-08 → dayDiff=7 → coldAllowed=false → cold=0, needCheck=true
  "TC15a": {
    testId: "TC15a",
    context: { patientId: "P001", treatDate: "2026-03-07",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "不全骨折", injuryDate: "2026-03-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  "TC15b": {
    testId: "TC15b",
    context: { patientId: "P001", treatDate: "2026-03-08",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "不全骨折", injuryDate: "2026-03-01", cold: true, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC16: 長期50%逓減（5か月超 + 月10回以上×5か月連続） ──────────────────
  // injuryDate=2026-02-01（day<16 → 2月起算）、treatDate=2026-07-15 → monthsElapsed=5
  // TC16a: monthlyVisitCounts=[10,10,10,10,10] → allFrequent=true → 50%
  // TC16b: monthlyVisitCounts=[10,10,9,10,10]  → month3=9<10 → 75%
  // TC16c: treatDate=2026-06-15 → monthsElapsed=4 → 100%（訪問頻度条件は関係なし）
  "TC16a": {
    testId: "TC16a",
    context: { patientId: "P001", treatDate: "2026-07-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true },
      monthlyVisitCounts: [10, 10, 10, 10, 10] },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "TC16b": {
    testId: "TC16b",
    context: { patientId: "P001", treatDate: "2026-07-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true },
      monthlyVisitCounts: [10, 10, 9, 10, 10] },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "TC16c": {
    testId: "TC16c",
    context: { patientId: "P001", treatDate: "2026-06-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true },
      monthlyVisitCounts: [10, 10, 10, 10, 10] },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC17: 温罨法 初検日特例 ────────────────────────────────────────────────
  // injuryDate=2026-03-11, treatDate=2026-03-17 → dayDiff=6（≥5 → 通常なら warm 可）
  // TC17a: kubun=初検 → 初検日特例で warm=0, needCheck=true
  // TC17b: kubun=後療 → 通常算定で warm=75, taiki=5
  "TC17a": {
    testId: "TC17a",
    context: { patientId: "P001", treatDate: "2026-03-17",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-03-11", cold: false, warm: true, electro: false }
      ]}
    ]
  },

  "TC17b": {
    testId: "TC17b",
    context: { patientId: "P001", treatDate: "2026-03-17",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-03-11", cold: false, warm: true, electro: false }
      ]}
    ]
  },

  // ── TC18: 長期継続理由書アラート ──────────────────────────────────────────
  // injuryDate=2026-02-01（day<16 → 2月起算）
  // TC18a: treatDate=2026-05-15 → monthsElapsed=3 → アラートあり
  // TC18b: treatDate=2026-04-15 → monthsElapsed=2 → アラートなし
  "TC18a": {
    testId: "TC18a",
    context: { patientId: "P001", treatDate: "2026-05-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  "TC18b": {
    testId: "TC18b",
    context: { patientId: "P001", treatDate: "2026-04-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-02-01", cold: false, warm: false, electro: false }
      ]}
    ]
  },

  // ── TC19: 金属副子等加算 Phase 1（§18.3）─────────────────────────────────
  // TC19a: 骨折 metalChk=true → metalOut=1000, 逓減対象外, needCheck=false
  // TC19b: 捻挫 metalChk=true → metalOut=0 + 要確認, needCheck=true
  "TC19a": {
    testId: "TC19a",
    context: { patientId: "P001", treatDate: "2026-01-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: true }
      ]}
    ]
  },

  "TC19b": {
    testId: "TC19b",
    context: { patientId: "P001", treatDate: "2026-01-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "腰部", byomei: "捻挫", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: true }
      ]}
    ]
  },

  // ── TC20: 金属副子等加算 Phase 2（caseKey 通算3回制限）─────────────────────
  // TC20a: 骨折 metalChk=true, metalPriorCount=0 → 1回目・算定可
  // TC20b: 骨折 metalChk=true, metalPriorCount=2 → 3回目・算定可
  // TC20c: 骨折 metalChk=true, metalPriorCount=3 → 上限超・算定不可
  "TC20a": {
    testId: "TC20a",
    context: { patientId: "P001", treatDate: "2026-01-15",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: true, metalPriorCount: 0 }
      ]}
    ]
  },

  "TC20b": {
    testId: "TC20b",
    context: { patientId: "P001", treatDate: "2026-01-25",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: true, metalPriorCount: 2 }
      ]}
    ]
  },

  "TC20c": {
    testId: "TC20c",
    context: { patientId: "P001", treatDate: "2026-02-05",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: true, metalPriorCount: 3 }
      ]}
    ]
  },

  // ── TC21: 柔道整復運動後療料（骨折・不全骨折・脱臼 / dayDiff >= 15）───────────────
  // TC21a: 骨折 dayDiff=15 exercise=true → exerciseOut=320, rowTotalOut=1170
  // TC21b: 骨折 dayDiff=14 exercise=true → 算定不可（15日未満）, rowTotalOut=850
  // TC21c: 捻挫 dayDiff=20 exercise=true → 算定不可（対象外傷病）, rowTotalOut=505
  // TC21d: 骨折 dayDiff=8  exercise=true → 算定不可（15日未満）, rowTotalOut=850
  "TC21a": {
    testId: "TC21a",
    context: { patientId: "P001", treatDate: "2026-01-20",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: false, exercise: true }
      ]}
    ]
  },

  "TC21b": {
    testId: "TC21b",
    context: { patientId: "P001", treatDate: "2026-01-19",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: false, exercise: true }
      ]}
    ]
  },

  "TC21c": {
    testId: "TC21c",
    context: { patientId: "P001", treatDate: "2026-01-25",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "捻挫", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: false, exercise: true }
      ]}
    ]
  },

  "TC21d": {
    testId: "TC21d",
    context: { patientId: "P001", treatDate: "2026-01-28",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-20",
          cold: false, warm: false, electro: false, metal: false, exercise: true }
      ]}
    ]
  },

  // ── TC22: 柔道整復運動後療料 Phase 2（当月5回上限）────────────────────────────
  // TC22a: 骨折 dayDiff=15 exercisePriorCount=5 → 算定上限超
  // TC22b: 骨折 dayDiff=15 exercisePriorCount=4 → 5回目・算定可
  "TC22a": {
    testId: "TC22a",
    context: { patientId: "P001", treatDate: "2026-01-20",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: false,
          exercise: true, exercisePriorCount: 5 }
      ]}
    ]
  },

  "TC22b": {
    testId: "TC22b",
    context: { patientId: "P001", treatDate: "2026-01-20",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "右前腕", byomei: "骨折", injuryDate: "2026-01-05",
          cold: false, warm: false, electro: false, metal: false,
          exercise: true, exercisePriorCount: 4 }
      ]}
    ]
  },

  // ── TC23: 特殊骨折初検 安全弁（未定義部位 → base=0）───────────────────────────
  // TC23a: 骨折 膝蓋骨（特殊骨折：mapBuiToSettingKey_ 未登録）→ base=0 + 要確認
  // TC23b: 骨折 腰椎（脊椎：mapBuiToSettingKey_ 未登録）→ base=0 + 要確認
  // TC23c: 骨折 胸骨（特殊骨折：mapBuiToSettingKey_ 未登録）→ base=0 + 要確認
  // TC23d: 骨折 大腿（定義済み部位）→ base=11800 正常算定（ポジティブケース）
  "TC23a": {
    testId: "TC23a",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "膝蓋骨", byomei: "骨折", injuryDate: "2026-03-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },
  "TC23b": {
    testId: "TC23b",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "腰椎", byomei: "骨折", injuryDate: "2026-03-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },
  "TC23c": {
    testId: "TC23c",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "胸骨", byomei: "骨折", injuryDate: "2026-03-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },
  "TC23d": {
    testId: "TC23d",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "大腿", byomei: "骨折", injuryDate: "2026-03-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },

  // ── TC24: 骨折/不全骨折 算定特性確認 ─────────────────────────────────────────
  // TC24a: 不全骨折 肩甲骨（未定義部位）→ base=0 + 要確認
  // TC24b: 骨折 後療 monthsElapsed=14 → ltCoef=1.0（§11 長期減額対象外）
  // TC24c: 不全骨折 後療 monthsElapsed=14 → ltCoef=1.0（§11 同上）
  "TC24a": {
    testId: "TC24a",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: false, reBilled: false, supportBilled: false } },
    cases: [
      { caseNo: 1, kubun: "初検", parts: [
        { bui: "肩甲骨", byomei: "不全骨折", injuryDate: "2026-03-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },
  "TC24b": {
    testId: "TC24b",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true },
      monthlyVisitCounts: [10, 10, 10, 10, 10] },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "前腕", byomei: "骨折", injuryDate: "2025-01-05",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },
  "TC24c": {
    testId: "TC24c",
    context: { patientId: "P001", treatDate: "2026-03-01",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true },
      monthlyVisitCounts: [10, 10, 10, 10, 10] },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "前腕", byomei: "不全骨折", injuryDate: "2025-01-05",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },

  // ── TC25: 脱臼長期逓減・骨折継続理由書アラートなし確認 ──────────────────────────
  // TC25a: 脱臼 後療 monthsElapsed=5 → ltCoef=0.75 + 継続理由書アラート（§18.2 脱臼対象）
  // TC25b: 骨折 後療 monthsElapsed=3 → ltCoef=1.0 + 継続理由書アラートなし（§20 骨折対象外・コード修正後）
  "TC25a": {
    testId: "TC25a",
    context: { patientId: "P001", treatDate: "2026-07-01",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "肩関節", byomei: "脱臼", injuryDate: "2026-02-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },
  "TC25b": {
    testId: "TC25b",
    context: { patientId: "P001", treatDate: "2026-05-01",
      monthlyStatus: { initBilled: true, reBilled: true, supportBilled: true } },
    cases: [
      { caseNo: 1, kubun: "後療", parts: [
        { bui: "前腕", byomei: "骨折", injuryDate: "2026-02-01",
          cold: false, warm: false, electro: false, metal: false, exercise: false }
      ]}
    ]
  },

};


/* =======================================================================
   Expected データ（tests/jrec01/expected/*.json と同内容）
   ======================================================================= */
var JREC01_EXPECTED_ = {

  "TC01": {
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 845, visitTotal: 2495,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-02-03_C1_P1", kubun: "初検", baseOut: 760, coldOut: 85, rowTotalOut: 845 }
    ]
  },

  "TC02": {
    header: { initFee: 0, reFee: 410, supportFee: 0, detailSum: 505, visitTotal: 915,
      needCheck: false, needCheckReason: "",
      billedKubun: "再検", mixedFlag: "通常",
      case1Summary: "case1:再検", case2Summary: "case2:なし", chargeReason: "再検のみ" },
    details: [
      { detailID: "P001_2026-02-05_C1_P1", kubun: "再検", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC03": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-07_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "M01": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1010, visitTotal: 1010,
      needCheck: true, needCheckReason: "同月別ケース初回 初検抑制",
      billedKubun: "後療", mixedFlag: "Mixed",
      case1Summary: "case1:後療", case2Summary: "case2:初検(抑制)", chargeReason: "初検抑制かつ再検対象なし" },
    details: [
      { detailID: "P001_2026-02-10_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-10_C2_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "M02": {
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 1265, visitTotal: 2915,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "Mixed",
      case1Summary: "case1:再検", case2Summary: "case2:初検", chargeReason: "算定可能な初検ありのため初検採用" },
    details: [
      { detailID: "P001_2026-02-03_C1_P1", kubun: "再検", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-03_C2_P1", kubun: "初検", baseOut: 760, coldOut: 0, rowTotalOut: 760 }
    ]
  },

  "M03": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1010, visitTotal: 1010,
      needCheck: true, needCheckReason: "同月別ケース初回 初検抑制",
      billedKubun: "後療", mixedFlag: "Mixed",
      case1Summary: "case1:後療", case2Summary: "case2:初検(抑制)", chargeReason: "初検抑制かつ再検対象なし" },
    details: [
      { detailID: "P001_2026-02-10_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-10_C2_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "M04": {
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 1520, visitTotal: 3170,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "Mixed",
      case1Summary: "case1:初検", case2Summary: "case2:初検", chargeReason: "算定可能な初検ありのため初検採用" },
    details: [
      { detailID: "P001_2026-02-10_C1_P1", kubun: "初検", baseOut: 760, coldOut: 0, rowTotalOut: 760 },
      { detailID: "P001_2026-02-10_C2_P1", kubun: "初検", baseOut: 760, coldOut: 0, rowTotalOut: 760 }
    ]
  },

  "M05": {
    header: { initFee: 0, reFee: 410, supportFee: 0, detailSum: 1010, visitTotal: 1420,
      needCheck: false, needCheckReason: "",
      billedKubun: "再検", mixedFlag: "Mixed",
      case1Summary: "case1:後療", case2Summary: "case2:再検", chargeReason: "再検ありのため再検採用" },
    details: [
      { detailID: "P001_2026-02-12_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-12_C2_P1", kubun: "再検", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // ── M06b: 治癒後別負傷（case2初検、case1は治癒済）──────────────────────────
  // initBilled=false（治癒後別負傷）→ initFee=1550
  // reBilled=true（case1再検算定済）→ reFee=0（hasReexam=false なので reBilled 無関係）
  // supportBilled=true（case1算定済）→ supportFee=0
  // [A] 施術継続中の再検抑制シナリオは TC09b 参照（2026-03-18 修正済）
  "M06b": {
    header: { initFee: 1550, reFee: 0, supportFee: 0, detailSum: 760, visitTotal: 2310,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-02-15_C1_P1", kubun: "初検", baseOut: 760, coldOut: 0, rowTotalOut: 760 }
    ]
  },

  // ── TC04 ──────────────────────────────────────────────────────────────
  "TC04a": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-04_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC04b": {
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 760, visitTotal: 2410,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-02-05_C1_P1", kubun: "初検", baseOut: 760, coldOut: 0, rowTotalOut: 760 }
    ]
  },

  // ── TC05 ──────────────────────────────────────────────────────────────
  "TC05a": {
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 845, visitTotal: 2495,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-02-01_C1_P1", kubun: "初検", baseOut: 760, coldOut: 85, rowTotalOut: 845 }
    ]
  },

  "TC05b": {
    header: { initFee: 0, reFee: 410, supportFee: 0, detailSum: 590, visitTotal: 1000,
      needCheck: false, needCheckReason: "",
      billedKubun: "再検", mixedFlag: "通常",
      case1Summary: "case1:再検", case2Summary: "case2:なし", chargeReason: "再検のみ" },
    details: [
      { detailID: "P001_2026-02-02_C1_P1", kubun: "再検", baseOut: 505, coldOut: 85, rowTotalOut: 590 }
    ]
  },

  "TC05c": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "冷罨法 算定不可（捻挫：受傷後2日）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-03_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // ── TC06 ──────────────────────────────────────────────────────────────
  "TC06a": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "温罨法 算定不可（捻挫：受傷後4日）;電療 算定不可（捻挫：受傷後4日）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-05_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, warmOut: 0, electroOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC06b": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 618, visitTotal: 618,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-06_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, warmOut: 75, electroOut: 33, rowTotalOut: 618 }
    ]
  },

  // ── TC07 ──────────────────────────────────────────────────────────────
  "TC07a": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: true, needCheckReason: "温罨法 算定不可（骨折：受傷後6日）;電療 算定不可（骨折：受傷後6日）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-07_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, warmOut: 0, electroOut: 0, rowTotalOut: 850 }
    ]
  },

  "TC07b": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 963, visitTotal: 963,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-08_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, warmOut: 75, electroOut: 33, rowTotalOut: 963 }
    ]
  },

  // ── TC08 ──────────────────────────────────────────────────────────────
  "TC08a": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 805, visitTotal: 805,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-05_C1_P1", kubun: "後療", baseOut: 720, coldOut: 85, rowTotalOut: 805 }
    ]
  },

  "TC08b": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 720, visitTotal: 720,
      needCheck: true, needCheckReason: "冷罨法 算定不可（脱臼：受傷後5日）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-06_C1_P1", kubun: "後療", baseOut: 720, coldOut: 0, rowTotalOut: 720 }
    ]
  },

  // ── TC09 ──────────────────────────────────────────────────────────────
  "TC09": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1010, visitTotal: 1010,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "Mixed",
      case1Summary: "case1:後療", case2Summary: "case2:後療", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-12_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-12_C2_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // ── TC09b: [A]施術継続中・case2再検抑制 ──────────────────────────────────
  // reBilled=true → reFee=0（2026-03-18 修正で正しく抑制）
  // billedKubun=後療: detailSum=505+505=1010
  "TC09b": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1010, visitTotal: 1010,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "Mixed",
      case1Summary: "case1:後療", case2Summary: "case2:再検", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-12_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-12_C2_P1", kubun: "再検", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // ── TC10 ──────────────────────────────────────────────────────────────
  "TC10": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1123, visitTotal: 1123,
      needCheck: true, needCheckReason: "同月別ケース初回 初検抑制;冷罨法 算定不可（捻挫：受傷後9日）",
      billedKubun: "後療", mixedFlag: "Mixed",
      case1Summary: "case1:後療", case2Summary: "case2:初検(抑制)", chargeReason: "初検抑制かつ再検対象なし" },
    details: [
      { detailID: "P001_2026-02-10_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-02-10_C2_P1", kubun: "後療", baseOut: 505, coldOut: 0, warmOut: 75, electroOut: 33, rowTotalOut: 618 }
    ]
  },

  // ── TC11 ──────────────────────────────────────────────────────────────
  "TC11": {
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 5200, visitTotal: 6850,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-02-01_C1_P1", kubun: "初検", baseOut: 5200, coldOut: 0, rowTotalOut: 5200 }
    ]
  },

  // ── TC12 ──────────────────────────────────────────────────────────────
  "TC12": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1010, visitTotal: 1010,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-03-01_C1_P2", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // ── TC13 ──────────────────────────────────────────────────────────────
  // 3部位目: 505 * 0.6 = 303（Math.round）
  "TC13": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1313, visitTotal: 1313,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-02_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-03-02_C1_P2", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 },
      { detailID: "P001_2026-03-02_C1_P3", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 303 }
    ]
  },

  // ── TC14 ──────────────────────────────────────────────────────────────
  // 長期逓減は月数ベース。baseOut は生値（505のまま）、rowTotalOut に ltCoef が反映される。
  "TC14a": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-06-30_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // TC14b: 505 * 0.75 = 378.75 → Math.round = 379
  "TC14b": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 379, visitTotal: 379,
      needCheck: true, needCheckReason: "長期減額75%適用（捻挫）;長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-07-01_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 379 }
    ]
  },

  // ── TC15: 不全骨折冷罨法 dayDiff境界 ──────────────────────────────────────
  // base=720（koryoFuzenKossetu）、cold=85（dayDiff<=6で許可）
  // TC15a: dayDiff=6 → coldAllowed=true → cold=85, rowTotalOut=805
  // TC15b: dayDiff=7 → coldAllowed=false → cold=0, rowTotalOut=720, needCheck=true
  "TC15a": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 805, visitTotal: 805,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-07_C1_P1", kubun: "後療", baseOut: 720, coldOut: 85, rowTotalOut: 805 }
    ]
  },

  "TC15b": {
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 720, visitTotal: 720,
      needCheck: true, needCheckReason: "冷罨法 算定不可（不全骨折：受傷後7日）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-08_C1_P1", kubun: "後療", baseOut: 720, coldOut: 0, rowTotalOut: 720 }
    ]
  },

  // ── TC16: 長期50%逓減 ──────────────────────────────────────────────────────
  // Math.round(505 * 0.50) = Math.round(252.5) = 253
  // Math.round(505 * 0.75) = Math.round(378.75) = 379
  "TC16a": {
    // 50%適用: monthsElapsed=5 かつ全月10回以上 → ltCoef=0.50 → 253
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 253, visitTotal: 253,
      needCheck: true, needCheckReason: "長期減額50%適用（捻挫）;長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-07-15_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 253 }
    ]
  },

  "TC16b": {
    // 75%のまま: monthsElapsed=5 だが月3=9回<10 → ltCoef=0.75 → 379
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 379, visitTotal: 379,
      needCheck: true, needCheckReason: "長期減額75%適用（捻挫）;長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-07-15_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 379 }
    ]
  },

  "TC16c": {
    // 減額なし: monthsElapsed=4（長期条件未達）→ ltCoef=1.0 → 505、継続理由書あり
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-06-15_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  // ── TC17: 温罨法 初検日特例 ────────────────────────────────────────────────
  // dayDiff=6（≥5 → 通常なら warm=75）。TC17a は初検日特例で warm=0。
  // shoryoNenZa=760, initFee=1550, supportFee=100 → visitTotal=2410
  "TC17a": {
    // 初検日 + warm要求 → 初検日特例で warm=0, needCheck=true
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 760, visitTotal: 2410,
      needCheck: true, needCheckReason: "温罨法 算定不可（初検日特例：捻挫）",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-03-17_C1_P1", kubun: "初検", baseOut: 760, coldOut: 0, rowTotalOut: 760 }
    ]
  },

  "TC17b": {
    // 後療日 + dayDiff=6 → 通常算定 warm=75, taiki=5 → rowTotalOut=585
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 585, visitTotal: 585,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-17_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 585 }
    ]
  },

  // ── TC18: 長期継続理由書アラート ──────────────────────────────────────────
  "TC18a": {
    // monthsElapsed=3 → アラートあり。ltCoef=1.0（減額なし）。base=505, rowTotalOut=505
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-05-15_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC18b": {
    // monthsElapsed=2 → アラートなし。needCheck=false
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-04-15_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC19a": {
    // 骨折 metalChk=true → metalOut=1000, rowTotalOut=850+1000=1850, needCheck=false
    // koryoKossetu=850, metalAddon=1000, coef=1.0, ltCoef=1.0(骨折は対象外)
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1850, visitTotal: 1850,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-15_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, metalOut: 1000, rowTotalOut: 1850 }
    ]
  },

  "TC19b": {
    // 捻挫 metalChk=true → metalOut=0 + 要確認, rowTotalOut=505, needCheck=true
    // koryoNenZa=505, coef=1.0, ltCoef=1.0(monthsElapsed=0)
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "金属副子等加算 算定不可（対象外傷病：捻挫）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-15_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, metalOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC20a": {
    // 骨折 metalPriorCount=0 → 1回目算定可。koryoKossetu=850, metalOut=1000, rowTotalOut=1850
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1850, visitTotal: 1850,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-15_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, metalOut: 1000, rowTotalOut: 1850 }
    ]
  },

  "TC20b": {
    // 骨折 metalPriorCount=2 → 3回目・算定可。koryoKossetu=850, metalOut=1000, rowTotalOut=1850
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1850, visitTotal: 1850,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-25_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, metalOut: 1000, rowTotalOut: 1850 }
    ]
  },

  "TC20c": {
    // 骨折 metalPriorCount=3 → 上限超・算定不可。koryoKossetu=850, metalOut=0, rowTotalOut=850
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: true, needCheckReason: "金属副子等加算 算定上限超（通算3回）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-02-05_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, metalOut: 0, rowTotalOut: 850 }
    ]
  },

  // ── TC21: 柔道整復運動後療料 ─────────────────────────────────────────────────
  "TC21a": {
    // 骨折 dayDiff=15 → 算定可。koryoKossetu=850, exerciseOut=320, rowTotalOut=1170
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1170, visitTotal: 1170,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-20_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, exerciseOut: 320, rowTotalOut: 1170 }
    ]
  },

  "TC21b": {
    // 骨折 dayDiff=14 → 算定不可（15日未満）。koryoKossetu=850, exerciseOut=0, rowTotalOut=850
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: true, needCheckReason: "運動後療料 算定不可（受傷後15日未満）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-19_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, exerciseOut: 0, rowTotalOut: 850 }
    ]
  },

  "TC21c": {
    // 捻挫 dayDiff=20 → 算定不可（対象外傷病）。koryoNenZa=505, exerciseOut=0, rowTotalOut=505
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 505, visitTotal: 505,
      needCheck: true, needCheckReason: "運動後療料 算定不可（対象外傷病：捻挫）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-25_C1_P1", kubun: "後療", baseOut: 505, coldOut: 0, exerciseOut: 0, rowTotalOut: 505 }
    ]
  },

  "TC21d": {
    // 骨折 dayDiff=8 → 算定不可（15日未満）。koryoKossetu=850, exerciseOut=0, rowTotalOut=850
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: true, needCheckReason: "運動後療料 算定不可（受傷後15日未満）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-28_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, exerciseOut: 0, rowTotalOut: 850 }
    ]
  },

  // ── TC22: 柔道整復運動後療料 Phase 2（当月5回上限）────────────────────────────
  "TC22a": {
    // 骨折 dayDiff=15 exercisePriorCount=5 → 上限超。koryoKossetu=850, exerciseOut=0, rowTotalOut=850
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: true, needCheckReason: "運動後療料 算定上限超（当月5回）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-20_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, exerciseOut: 0, rowTotalOut: 850 }
    ]
  },

  "TC22b": {
    // 骨折 dayDiff=15 exercisePriorCount=4 → 5回目・算定可。koryoKossetu=850, exerciseOut=320, rowTotalOut=1170
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 1170, visitTotal: 1170,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-01-20_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, exerciseOut: 320, rowTotalOut: 1170 }
    ]
  },

  // ── TC23: 特殊骨折初検 安全弁 ──────────────────────────────────────────────
  "TC23a": {
    // 骨折 膝蓋骨（mapBuiToSettingKey_ 未登録）→ base=0, needCheck=true
    // visitTotal = initFee(1550) + supportFee(100) + detailSum(0) = 1650
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 0, visitTotal: 1650,
      needCheck: true,
      needCheckReason: "整復料/固定料 取得不可（膝蓋骨：設定シートにキーがありません）",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "初検", baseOut: 0, coldOut: 0, rowTotalOut: 0 }
    ]
  },
  "TC23b": {
    // 骨折 腰椎（脊椎：mapBuiToSettingKey_ 未登録）→ base=0, needCheck=true
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 0, visitTotal: 1650,
      needCheck: true,
      needCheckReason: "整復料/固定料 取得不可（腰椎：設定シートにキーがありません）",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "初検", baseOut: 0, coldOut: 0, rowTotalOut: 0 }
    ]
  },
  "TC23c": {
    // 骨折 胸骨（特殊骨折：mapBuiToSettingKey_ 未登録）→ base=0, needCheck=true
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 0, visitTotal: 1650,
      needCheck: true,
      needCheckReason: "整復料/固定料 取得不可（胸骨：設定シートにキーがありません）",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "初検", baseOut: 0, coldOut: 0, rowTotalOut: 0 }
    ]
  },
  "TC23d": {
    // 骨折 大腿（定義済み部位）→ base=整復料_骨折_大腿=11800, needCheck=false（ポジティブケース）
    // visitTotal = initFee(1550) + supportFee(100) + detailSum(11800) = 13450
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 11800, visitTotal: 13450,
      needCheck: false, needCheckReason: "",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "初検", baseOut: 11800, coldOut: 0, rowTotalOut: 11800 }
    ]
  },

  // ── TC24: 骨折/不全骨折 算定特性確認 ─────────────────────────────────────────
  "TC24a": {
    // 不全骨折 肩甲骨（mapBuiToSettingKey_ 未登録）→ base=0, needCheck=true
    header: { initFee: 1550, reFee: 0, supportFee: 100, detailSum: 0, visitTotal: 1650,
      needCheck: true,
      needCheckReason: "整復料/固定料 取得不可（肩甲骨：設定シートにキーがありません）",
      billedKubun: "初検", mixedFlag: "通常",
      case1Summary: "case1:初検", case2Summary: "case2:なし", chargeReason: "初検のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "初検", baseOut: 0, coldOut: 0, rowTotalOut: 0 }
    ]
  },
  "TC24b": {
    // 骨折 後療 monthsElapsed=14 → ltCoef=1.0（§11 骨折は長期減額対象外）+ 継続理由書アラートなし（§20 骨折対象外）
    // koryoKossetu=850, rowTotalOut=850, needCheck=false
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, rowTotalOut: 850 }
    ]
  },
  "TC24c": {
    // 不全骨折 後療 monthsElapsed=14 → ltCoef=1.0（§11 不全骨折も長期減額対象外）+ 継続理由書アラートなし
    // koryoFuzenKossetu=720, rowTotalOut=720, needCheck=false
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 720, visitTotal: 720,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-03-01_C1_P1", kubun: "後療", baseOut: 720, coldOut: 0, rowTotalOut: 720 }
    ]
  },

  // ── TC25: 脱臼長期逓減・骨折継続理由書アラートなし ─────────────────────────────
  "TC25a": {
    // 脱臼 後療 monthsElapsed=5 → ltCoef=0.75（§11 脱臼は対象）
    // koryoDakkyu=720, rowTotalOut=Math.round(720*0.75)=540, needCheck=true（長期75%+継続理由書）
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 540, visitTotal: 540,
      needCheck: true,
      needCheckReason: "長期減額75%適用（脱臼）;長期施術3か月超（継続理由書確認）",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-07-01_C1_P1", kubun: "後療", baseOut: 720, coldOut: 0, rowTotalOut: 540 }
    ]
  },
  "TC25b": {
    // 骨折 後療 monthsElapsed=3 → ltCoef=1.0（§11 骨折対象外）+ 継続理由書アラートなし（§20 骨折対象外・コード修正後）
    // koryoKossetu=850, rowTotalOut=850, needCheck=false
    header: { initFee: 0, reFee: 0, supportFee: 0, detailSum: 850, visitTotal: 850,
      needCheck: false, needCheckReason: "",
      billedKubun: "後療", mixedFlag: "通常",
      case1Summary: "case1:後療", case2Summary: "case2:なし", chargeReason: "後療のみ" },
    details: [
      { detailID: "P001_2026-05-01_C1_P1", kubun: "後療", baseOut: 850, coldOut: 0, rowTotalOut: 850 }
    ]
  },

};


/* =======================================================================
   computeAmountsFromFixture_V3_
   fixture → amounts 変換（calcHeaderAmountsByVisitKey_V3_ の純粋計算部分を複製）
   ※ production ロジック変更時はここも更新すること
   ======================================================================= */
function computeAmountsFromFixture_V3_(fx) {
  var settings             = TEST_SETTINGS_;
  var ms                   = fx.context.monthlyStatus;
  var treatDate            = new Date(fx.context.treatDate);
  var patId                = fx.context.patientId;
  var kubun1               = (fx.cases[0] || {}).kubun || null;
  var kubun2               = (fx.cases[1] || {}).kubun || null;
  var monthlyVisitCounts   = fx.context.monthlyVisitCounts || null;  // §11 50%逓減用
  var reasons              = [];

  // --- 初検料 ---
  var hasInit   = (kubun1 === "初検" || kubun2 === "初検");
  var hasReexam = (kubun1 === "再検" || kubun2 === "再検");
  var hasKoryo  = (kubun1 === "再検" || kubun1 === "後療" || kubun2 === "再検" || kubun2 === "後療");

  var initFee = 0;
  if (hasInit) {
    if (ms.initBilled) {
      reasons.push("同月別ケース初回 初検抑制");
    } else {
      initFee = settings.initFee;
    }
  }
  var hasBillableInitial = (initFee > 0);

  // --- 相談支援料 ---
  var supportFee = 0;
  if (hasBillableInitial) {
    supportFee = ms.supportBilled ? 0 : settings.initSupport;
  }

  // --- 再検料 ---
  // monthlyStatus.reBilled 制御（2026-03-18 追加, amounts.js と同期）:
  //   [A] 施術継続中: reBilled=true → 当月2件目の再検を抑制（reFee=0）
  //   [B] 治癒後別負傷: isCaseEndedBefore_ で suppressReBilled=true → reBilled=false → 再検許可
  var reFee = 0;
  if (hasReexam && !hasBillableInitial && !ms.reBilled) {
    reFee = settings.reFee;
  }

  // --- 実効区分（抑制変換） ---
  var calcKoryoOnThisDay = !hasBillableInitial;
  var effectiveKubun1 = calcKoryoOnThisDay ? (kubun1 === "初検" ? "後療" : kubun1) : kubun1;
  var effectiveKubun2 = calcKoryoOnThisDay ? (kubun2 === "初検" ? "後療" : kubun2) : kubun2;

  // --- 部位別明細計算 ---
  function calcPartsFromFixture_(caseData, effectiveKubun) {
    if (!caseData || !caseData.parts) return { total: 0, parts: [] };
    var total = 0;
    var parts = [];
    for (var i = 0; i < caseData.parts.length; i++) {
      var p = caseData.parts[i];
      var injDate = new Date(p.injuryDate);
      var part = calcOnePartAmount_V3_(
        settings, effectiveKubun, p.byomei, injDate, treatDate,
        !!p.cold, !!p.warm, !!p.electro,
        i + 1, reasons, p.bui, monthlyVisitCounts,
        !!p.metal,                                                                   // §18.3 Phase 1
        (p.metalPriorCount !== undefined) ? Number(p.metalPriorCount) : null,        // §18.3 Phase 2
        !!p.exercise,                                                                // 運動後療料 Phase 1
        (p.exercisePriorCount !== undefined) ? Number(p.exercisePriorCount) : null   // 運動後療料 Phase 2
      );
      part.bui = p.bui;
      total += part.total;
      parts.push(part);
    }
    return { total: total, parts: parts };
  }

  var detail1   = calcPartsFromFixture_(fx.cases[0], effectiveKubun1);
  var detail2   = calcPartsFromFixture_(fx.cases[1], effectiveKubun2);
  var detailSum = detail1.total + detail2.total;
  var visitTotal = initFee + reFee + supportFee + detailSum;

  // --- 新5列 ---
  var isMixed       = (kubun2 != null && String(kubun2).trim() !== "");
  var initSuppressed = reasons.some(function(r) { return r.indexOf("初検抑制") !== -1; });

  var billedKubun = initFee > 0 ? "初検" : reFee > 0 ? "再検" : hasKoryo ? "後療" : "算定なし";
  var mixedFlag   = isMixed ? "Mixed" : "通常";

  var k1 = String(kubun1 || "").trim();
  var case1Summary = k1 === "初検" ? "case1:初検"
    : k1 === "再検" ? "case1:再検"
    : k1 === "後療" ? "case1:後療"
    : "case1:なし";

  var k2 = String(kubun2 || "").trim();
  var case2Summary;
  if (!k2)             case2Summary = "case2:なし";
  else if (k2 === "初検") case2Summary = initSuppressed ? "case2:初検(抑制)" : "case2:初検";
  else if (k2 === "再検") case2Summary = "case2:再検";
  else if (k2 === "後療") case2Summary = "case2:後療";
  else                    case2Summary = "case2:" + k2;

  var chargeReason;
  if      (hasBillableInitial && !isMixed)                              chargeReason = "初検のみ";
  else if (hasBillableInitial && isMixed)                               chargeReason = "算定可能な初検ありのため初検採用";
  else if (!hasBillableInitial && reFee > 0 && isMixed && initSuppressed)  chargeReason = "初検抑制のため再検採用";
  else if (!hasBillableInitial && reFee > 0 && isMixed && !initSuppressed) chargeReason = "再検ありのため再検採用";
  else if (!hasBillableInitial && reFee > 0 && !isMixed)               chargeReason = "再検のみ";
  else if (!hasBillableInitial && reFee === 0 && hasKoryo && isMixed && initSuppressed) chargeReason = "初検抑制かつ再検対象なし";
  else if (!hasBillableInitial && reFee === 0 && hasKoryo)             chargeReason = "後療のみ";
  else                                                                   chargeReason = "算定なし";

  // --- detail リスト生成 ---
  var details = [];
  var pushDetails_ = function(parts, caseNo, ek) {
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      details.push({
        detailID:    patId + "_" + fx.context.treatDate + "_C" + caseNo + "_P" + (i + 1),
        kubun:       ek,
        baseOut:     p.base,
        coldOut:     p.cold,
        warmOut:     p.warm,
        electroOut:  p.electro,
        metalOut:    p.metalOut,     // §18.3
        exerciseOut: p.exerciseOut,  // 柔道整復運動後療料
        rowTotalOut: Math.round(p.total),
      });
    }
  };
  pushDetails_(detail1.parts, 1, effectiveKubun1);
  pushDetails_(detail2.parts, 2, effectiveKubun2);

  return {
    initFee: initFee, reFee: reFee, supportFee: supportFee,
    detailSum: detailSum, visitTotal: visitTotal,
    needCheck: reasons.length > 0,
    needCheckReason: reasons.join(";"),
    billedKubun: billedKubun, mixedFlag: mixedFlag,
    case1Summary: case1Summary, case2Summary: case2Summary,
    chargeReason: chargeReason,
    effectiveKubun1: effectiveKubun1, effectiveKubun2: effectiveKubun2,
    details: details,
  };
}


/* =======================================================================
   assertAmounts_  ―  actual vs expected 比較
   ======================================================================= */
function assertAmounts_(testId, actual, expected) {
  var diffs = [];

  // header 比較
  var hKeys = Object.keys(expected.header);
  for (var i = 0; i < hKeys.length; i++) {
    var k = hKeys[i];
    var a = actual[k];
    var e = expected.header[k];
    if (String(a) !== String(e)) {
      diffs.push("header." + k + ": expect=" + JSON.stringify(e) + " actual=" + JSON.stringify(a));
    }
  }

  // detail 比較
  var exDetails = expected.details || [];
  for (var j = 0; j < exDetails.length; j++) {
    var ex = exDetails[j];
    var ac = actual.details[j];
    if (!ac) {
      diffs.push("detail[" + j + "]: missing");
      continue;
    }
    var dKeys = Object.keys(ex);
    for (var m = 0; m < dKeys.length; m++) {
      var dk = dKeys[m];
      if (String(ac[dk]) !== String(ex[dk])) {
        diffs.push("detail[" + j + "]." + dk + ": expect=" + JSON.stringify(ex[dk]) + " actual=" + JSON.stringify(ac[dk]));
      }
    }
  }

  return { pass: diffs.length === 0, diff: diffs.join(" / ") };
}


/* =======================================================================
   runFixtureTest_  ―  個別 fixture 実行
   ======================================================================= */
function runFixtureTest_(testId) {
  var fx = JREC01_FIXTURES_[testId];
  var ex = JREC01_EXPECTED_[testId];
  if (!fx) throw new Error("fixture not found: " + testId);
  if (!ex) throw new Error("expected not found: " + testId);

  var result = computeAmountsFromFixture_V3_(fx);
  return assertAmounts_(testId, result, ex);
}


/* =======================================================================
   runFixtureSuite_  ―  全 fixture 一括実行（メニューから呼び出す）
   ======================================================================= */
function runFixtureSuite_() {
  var ids  = Object.keys(JREC01_FIXTURES_);
  var pass = 0, fail = 0;
  var log  = [];

  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    try {
      var r = runFixtureTest_(id);
      if (r.pass) {
        pass++;
        log.push("[PASS] " + id);
      } else {
        fail++;
        log.push("[FAIL] " + id + "\n       " + r.diff);
      }
    } catch (e) {
      fail++;
      log.push("[ERROR] " + id + "\n       " + e.message);
    }
  }

  var summary = "PASS: " + pass + "  FAIL: " + fail + "  / " + ids.length;
  Logger.log(summary + "\n\n" + log.join("\n"));
  SpreadsheetApp.getUi().alert(summary + "\n\n" + log.join("\n"));
}


/**
 * B-1 Web API: 全 fixture テストをスイート実行して JSON で返す
 * google.script.run 経由で live-check-runner から呼び出し可能。
 * SpreadsheetApp.getUi() を使わないため Web App コンテキストで動作する。
 *
 * @returns {{ ok, passCount, failCount, total, results: Array, summary: string }}
 */
function runFixtureSuiteWeb_V3() {
  try {
    var ids    = Object.keys(JREC01_FIXTURES_);
    var pass   = 0, fail = 0;
    var results = [];

    for (var i = 0; i < ids.length; i++) {
      var id = ids[i];
      try {
        var r = runFixtureTest_(id);
        if (r.pass) {
          pass++;
          results.push({ testId: id, pass: true, diff: "" });
        } else {
          fail++;
          results.push({ testId: id, pass: false, diff: String(r.diff || "") });
        }
      } catch (e) {
        fail++;
        results.push({ testId: id, pass: false, diff: "ERROR: " + e.message });
      }
    }

    var summary = "PASS: " + pass + "  FAIL: " + fail + "  / " + ids.length;
    Logger.log("[runFixtureSuiteWeb_V3] " + summary);
    return { ok: true, passCount: pass, failCount: fail, total: ids.length, results: results, summary: summary };

  } catch (e) {
    Logger.log("[runFixtureSuiteWeb_V3] error=" + e.message);
    return { ok: false, passCount: 0, failCount: 0, total: 0, results: [], summary: e.message };
  }
}

/* =======================================================================
   公開ラッパー関数（Apps Script 実行メニューに表示される）
   末尾アンダースコアなし・引数なし
   ======================================================================= */
function runFixtureSuite()  { runFixtureSuite_(); }
function runFixtureTC01()   { showFixtureResult_("TC01"); }
function runFixtureTC02()   { showFixtureResult_("TC02"); }
function runFixtureTC03()   { showFixtureResult_("TC03"); }
function runFixtureTC04a()  { showFixtureResult_("TC04a"); }
function runFixtureTC04b()  { showFixtureResult_("TC04b"); }
function runFixtureTC05a()  { showFixtureResult_("TC05a"); }
function runFixtureTC05b()  { showFixtureResult_("TC05b"); }
function runFixtureTC05c()  { showFixtureResult_("TC05c"); }
function runFixtureTC06a()  { showFixtureResult_("TC06a"); }
function runFixtureTC06b()  { showFixtureResult_("TC06b"); }
function runFixtureTC07a()  { showFixtureResult_("TC07a"); }
function runFixtureTC07b()  { showFixtureResult_("TC07b"); }
function runFixtureTC08a()  { showFixtureResult_("TC08a"); }
function runFixtureTC08b()  { showFixtureResult_("TC08b"); }
function runFixtureTC09()   { showFixtureResult_("TC09"); }
function runFixtureTC09b()  { showFixtureResult_("TC09b"); }
function runFixtureTC10()   { showFixtureResult_("TC10"); }
function runFixtureTC11()   { showFixtureResult_("TC11"); }
function runFixtureTC12()   { showFixtureResult_("TC12"); }
function runFixtureTC13()   { showFixtureResult_("TC13"); }
function runFixtureTC14a()  { showFixtureResult_("TC14a"); }
function runFixtureTC14b()  { showFixtureResult_("TC14b"); }
function runFixtureTC15a()  { showFixtureResult_("TC15a"); }
function runFixtureTC15b()  { showFixtureResult_("TC15b"); }
function runFixtureTC16a()  { showFixtureResult_("TC16a"); }
function runFixtureTC16b()  { showFixtureResult_("TC16b"); }
function runFixtureTC16c()  { showFixtureResult_("TC16c"); }
function runFixtureTC17a()  { showFixtureResult_("TC17a"); }
function runFixtureTC17b()  { showFixtureResult_("TC17b"); }
function runFixtureTC18a()  { showFixtureResult_("TC18a"); }
function runFixtureTC18b()  { showFixtureResult_("TC18b"); }
function runFixtureTC19a()  { showFixtureResult_("TC19a"); }
function runFixtureTC19b()  { showFixtureResult_("TC19b"); }
function runFixtureTC20a()  { showFixtureResult_("TC20a"); }
function runFixtureTC20b()  { showFixtureResult_("TC20b"); }
function runFixtureTC20c()  { showFixtureResult_("TC20c"); }
function runFixtureTC21a()  { showFixtureResult_("TC21a"); }
function runFixtureTC21b()  { showFixtureResult_("TC21b"); }
function runFixtureTC21c()  { showFixtureResult_("TC21c"); }
function runFixtureTC21d()  { showFixtureResult_("TC21d"); }
function runFixtureTC22a()  { showFixtureResult_("TC22a"); }
function runFixtureTC22b()  { showFixtureResult_("TC22b"); }
function runFixtureTC23a()  { showFixtureResult_("TC23a"); }
function runFixtureTC23b()  { showFixtureResult_("TC23b"); }
function runFixtureTC23c()  { showFixtureResult_("TC23c"); }
function runFixtureTC23d()  { showFixtureResult_("TC23d"); }
function runFixtureTC24a()  { showFixtureResult_("TC24a"); }
function runFixtureTC24b()  { showFixtureResult_("TC24b"); }
function runFixtureTC24c()  { showFixtureResult_("TC24c"); }
function runFixtureTC25a()  { showFixtureResult_("TC25a"); }
function runFixtureTC25b()  { showFixtureResult_("TC25b"); }
function runFixtureM01()    { showFixtureResult_("M01"); }
function runFixtureM02()    { showFixtureResult_("M02"); }
function runFixtureM03()    { showFixtureResult_("M03"); }
function runFixtureM04()    { showFixtureResult_("M04"); }
function runFixtureM05()    { showFixtureResult_("M05"); }
function runFixtureM06b()   { showFixtureResult_("M06b"); }

function showFixtureResult_(testId) {
  var r = runFixtureTest_(testId);
  var msg = r.pass
    ? "[PASS] " + testId
    : "[FAIL] " + testId + "\n\n" + r.diff;
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}


/* =======================================================================
   D2 継続月数・頻回 純粋計算テスト（シート不要）
   Apps Script エディタで runD2Suite() を実行して確認する。

   ★制度ルール（確認点）:
     ① 4連続+当月10回以上 → displayMonths=5（5か月目表示）
     ② 5連続達成翌月       → displayMonths=6（固定）
     ③ 頻回開始後に当月10回未満でも → displayMonths=6（継続）
     ④ 16日以降初検の翌月起算  → totalMonths が1か月短くなること
     ⑤ case2行はM31空欄        → コードレビューで確認済み（if caseNo===1）
   ======================================================================= */

/**
 * D2 純粋計算ロジック（V3TR_calcD2Keizoku_ の計算部のみを分離）
 * countsArr: 0-indexed 月インデックス → 月別来院日数 の配列（長さ = totalMonths）
 * totalMonths: 起算月〜対象月の月数
 */
function V3TR_calcD2FromCounts_(countsArr, totalMonths) {
  if (!countsArr || totalMonths <= 0) return { rawContMonths: 0, freqStarted: false, displayMonths: "" };

  var streak = 0;
  var freqStartedBefore = false;
  for (var m = 0; m < totalMonths - 1; m++) {
    var cnt = countsArr[m] || 0;
    if (cnt >= 10) {
      streak++;
      if (streak >= 5) { freqStartedBefore = true; break; }
    } else {
      streak = 0;
    }
  }
  if (freqStartedBefore) {
    return { rawContMonths: -1, freqStarted: true, displayMonths: 6 };
  }
  var curCnt = countsArr[totalMonths - 1] || 0;
  if (curCnt >= 10) {
    streak++;
  } else {
    streak = 0;
  }
  return {
    rawContMonths: streak,
    freqStarted:   (streak >= 5),
    displayMonths: (streak > 0) ? streak : "",
  };
}

/**
 * D2 起算月計算（V3TR_calcD2Keizoku_ の起算月ロジックを分離）
 * injuryDate: Date、ym: "yyyy-MM"
 * return: 起算月〜対象月の月数（totalMonths）
 */
function V3TR_calcD2TotalMonths_(injuryDate, ym) {
  var sy = injuryDate.getFullYear();
  var sm = injuryDate.getMonth();
  if (injuryDate.getDate() >= 16) {
    sm++;
    if (sm > 11) { sm = 0; sy++; }
  }
  var parts = ym.split("-").map(Number);
  var ey = parts[0], em = parts[1] - 1;
  return (ey - sy) * 12 + (em - sm) + 1;
}

var D2_TEST_CASES_ = [
  {
    id: "D2-①",
    desc: "4連続+当月10回以上 → displayMonths=5（5か月目）",
    // 起算月1〜4か月目: 全て≥10。対象月（5か月目）: ≥10
    counts:       [12, 10, 11, 10, 10],
    totalMonths:  5,
    expectedDisp: 5,
    expectedFreq: true,
  },
  {
    id: "D2-②",
    desc: "5連続達成翌月（6か月目） → displayMonths=6固定",
    // 起算月1〜5か月目: 全て≥10（freqStartedBeforeで検出）。対象月（6か月目）: 0回でも
    counts:       [12, 10, 11, 10, 10, 0],
    totalMonths:  6,
    expectedDisp: 6,
    expectedFreq: true,
  },
  {
    id: "D2-③",
    desc: "頻回開始後に当月10回未満でも → displayMonths=6継続",
    // 起算月1〜5: ≥10。対象月（7か月目）: 5回
    counts:       [12, 10, 11, 10, 10, 0, 5],
    totalMonths:  7,
    expectedDisp: 6,
    expectedFreq: true,
  },
  {
    id: "D2-④a",
    desc: "16日以降初検(2026-01-20) → 起算月=2026-02。対象月2026-06 → totalMonths=5",
    injuryDate:      new Date(2026, 0, 20),  // 2026-01-20（1月=0-indexed）
    ym:              "2026-06",
    expectedTotal:   5,  // 2026-02〜2026-06 = 5ヶ月
  },
  {
    id: "D2-④b",
    desc: "15日以前初検(2026-01-15) → 起算月=2026-01。対象月2026-06 → totalMonths=6",
    injuryDate:      new Date(2026, 0, 15),  // 2026-01-15
    ym:              "2026-06",
    expectedTotal:   6,  // 2026-01〜2026-06 = 6ヶ月
  },
  {
    id: "D2-⑤",
    desc: "連続途切れ後に再び5連続 → 2度目の5連続翌月にdisplayMonths=6",
    // 1〜3: ≥10、4: <10（途切れ）、5〜9: ≥10（5連続再達成）、10: 対象月
    counts:       [12, 10, 11, 8, 10, 11, 12, 10, 10, 0],
    totalMonths:  10,
    expectedDisp: 6,
    expectedFreq: true,
  },
];

function runD2Suite_() {
  var pass = 0, fail = 0;
  var log = [];

  for (var i = 0; i < D2_TEST_CASES_.length; i++) {
    var tc = D2_TEST_CASES_[i];
    var ok = false;
    var detail = "";

    if (tc.expectedTotal !== undefined) {
      // ④ 起算月テスト
      var total = V3TR_calcD2TotalMonths_(tc.injuryDate, tc.ym);
      ok = (total === tc.expectedTotal);
      detail = "totalMonths: " + total + " (expected " + tc.expectedTotal + ")";
    } else {
      // ①②③⑤ 計算ロジックテスト
      var res = V3TR_calcD2FromCounts_(tc.counts, tc.totalMonths);
      ok = (res.displayMonths === tc.expectedDisp && res.freqStarted === tc.expectedFreq);
      detail = "displayMonths=" + res.displayMonths + "(expect " + tc.expectedDisp + ")"
        + " freqStarted=" + res.freqStarted + "(expect " + tc.expectedFreq + ")";
    }

    if (ok) {
      pass++;
      log.push("[PASS] " + tc.id + ": " + tc.desc);
    } else {
      fail++;
      log.push("[FAIL] " + tc.id + ": " + tc.desc + "\n       " + detail);
    }
  }

  var summary = "D2テスト PASS: " + pass + "  FAIL: " + fail + "  / " + D2_TEST_CASES_.length;
  Logger.log(summary + "\n" + log.join("\n"));
  SpreadsheetApp.getUi().alert(summary + "\n\n" + log.join("\n"));
}

/** ⑤ case2のM31空欄確認（コードレビュー用メモ）
 * V3TR_buildTransferDataForMonth_ 内の for(caseNo of [1,2]) ループで
 * caseNo===1 のブロックのみ V3TR_calcD2Keizoku_ を呼び出し row["経過"] を設定。
 * caseNo===2 は else { row["経過"] = ""; } で空文字を設定している。
 * → write_application.py では `if keizoku:` の条件で空文字はスキップされ M31 未書込になる。
 */

// 公開ラッパー（Args Scriptメニューから直接実行可能）
function runD2Suite() { runD2Suite_(); }
