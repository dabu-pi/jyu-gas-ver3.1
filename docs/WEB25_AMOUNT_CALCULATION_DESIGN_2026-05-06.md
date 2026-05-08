# WEB-2.5 設計書 — Web UI 来院登録 × 金額算定連携

作成日: 2026-05-06  
担当: dabu-pi  
前提フェーズ: WEB-2 完了（saveVisitFromWeb_V3 / web-visit-new.html 実装・LiveCheck 42 PASS）  
ブランチ: `feature/auto-dev-phase3-loop`

---

## 1. 目的と方針

### 目的

WEB-2 で実装した `saveVisitFromWeb_V3` は、来院ケース・来院ヘッダに保存するが **金額は0 / 要確認=true** で暫定保存していた。  
WEB-2.5 では、既存の保険算定ロジック（`calcHeaderAmountsByVisitKey_V3_`）を Web 版から呼び出し、  
**正確な候補金額をヘッダに記録し、要確認理由も明記**する。

### 設計原則

| 原則 | 内容 |
|---|---|
| レセプト事故ゼロ | Web 登録は常に `needCheck=true`。請求確定は人間がスプレッドシートで確認後に行う |
| 算定事実主義 | 既存の `calcHeaderAmountsByVisitKey_V3_` と同じロジックで算定候補を計算する |
| 抑制時は理由を残す | `needCheckReason` に "Web UI 登録" + 抑制理由をセミコロン区切りで記録 |
| 既存 UI 非破壊 | `saveVisit_V3`（スプレッドシートUI保存）は一切変更しない |
| 段階的移行 | 施術明細・初検情報履歴は WEB-2.5 MVP 対象外（WEB-2.5.1 以降） |

---

## 2. 既存ロジック調査結果

### 2-1. `calcHeaderAmountsByVisitKey_V3_`（Ver3_amounts.js:658）

**シグネチャ:**
```javascript
calcHeaderAmountsByVisitKey_V3_(ss, visitKey, patientId, treatDate, kubun1, kubun2)
```

**読み取りシート一覧:**

| シート | 目的 | 書き込みあり |
|---|---|---|
| 設定 | 単価・負担割合単位 | なし |
| 患者マスタ | 負担割合 | なし |
| 来院ヘッダ | 月内算定状況（30日ルール） | なし |
| 来院ケース | 部位・傷病・受傷日・冷/温/電 | なし |
| 施術明細 | 金属副子/運動後療の通算回数 | なし |

→ **全て読み取りのみ。SpreadsheetApp.getUi() 呼び出しなし。Web版から呼べる。**

**前提条件:**  
来院ケースに対象の `visitKey + caseNo` の行が保存済みであること（WEB-2 でケースを先に保存するため満たされる）。

**返却値（主要）:**

```javascript
{
  initFee, reFee, supportFee, detailSum, visitTotal,
  windowPay, claimPay,
  needCheck,           // 抑制理由があれば true
  needCheckReason,     // 抑制理由（"同月別ケース初回 初検抑制" 等）
  billedKubun,         // 課金実績の代表区分
  mixedFlag,           // "Mixed" or "通常"
  case1Summary, case2Summary, chargeReason,  // 説明性列
  details: { case1Parts, case2Parts }        // 部位別金額詳細
}
```

### 2-2. `calcEpisodeForCase_`（Ver3_core.js:3345）

**読み取りシート:** 来院ケースのみ（書き込みなし）  
**判定ロジック:**  
- 直近30日ルール → エピソード連続性
- 来院回数（0回: 初検、1回: 再検、2回以上: 後療）
- 治癒（endDate + 転帰）チェック

→ **UI シート非依存。Web 版から呼べる。`saveVisit_V3` と同一ロジック。**

### 2-3. `upsertOneCase_`（Ver3_core.js:1619）

**UI シート依存（Web 版から直接呼べない）:**
- `uiSh`（患者画面シート）から部位・傷病・受傷日・所見 を読み取る
- `readRowNewUI_` / `getMergedValue_` 呼び出し

→ WEB-2 でケース保存には専用のコードを使用済み（payload から直接書き込み）。

### 2-4. `appendInitHistory_V3_`（Ver3_core.js:4143）

**書き込み先:** 初検情報履歴シート  
**必要データ:** 負傷の日時・場所・状況・初検時所見 など（`readInitInfoFromUI_` で UIシートから取得）

→ **WEB-2.5 MVP スコープ外。**  
Web フォームに 負傷詳細フィールドがないため省略。  
needCheckReason に "初検情報履歴未記録" を追加することで、スプレッドシートUIでの後記録を促す。

### 2-5. `upsertDetailRows_V3_`（Ver3_core.js:1532）

**書き込み先:** 施術明細シート  
**必要データ:** `amounts.details`（calcHeaderAmountsByVisitKey_V3_ の返却値）+ `ep1.episodeStartDate` / `ep2.episodeStartDate`

→ **WEB-2.5 MVP スコープ外（WEB-2.5.1 で対応）。**  
施術明細なしの来院ヘッダは `recalcAmountsByVisitKey_V3_` で施術明細から再計算できないが、  
スプレッドシートUIから `saveVisit_V3` を呼べば上書きされる（既存運用で解消可能）。

### 2-6. `appendHeaderRow_V3_` vs ヘッダ更新関数の有無

`updateHeaderRow_V3_` は現時点で存在しない。  
→ WEB-2.5 では `saveVisitFromWeb_V3` 内で **ケース保存 → 算定 → ヘッダ保存** の順にし、  
ゼロ金額でヘッダを先保存せず、**算定後の実金額で1回だけ `appendHeaderRow_V3_` を呼ぶ**。

---

## 3. WEB-2.5 の実装範囲（MVP）

### 対象範囲（IN SCOPE）

| # | 内容 |
|---|---|
| 1 | `saveVisitFromWeb_V3` 改修: kubun を `calcEpisodeForCase_` で自動判定 |
| 2 | `saveVisitFromWeb_V3` 改修: `calcHeaderAmountsByVisitKey_V3_` 呼び出し |
| 3 | 来院ヘッダに算定候補金額を保存（`needCheck=true` / 理由付き） |
| 4 | `web-visit-new.html` に算定候補金額の表示（保存成功後） |
| 5 | docs / PROJECT_STATUS.md 更新 |
| 6 | LiveCheck（W2-6/7 の確認モーダルに金額表示を確認） |

### スコープ外（WEB-2.5 以降）

| # | 内容 | 理由 |
|---|---|---|
| 7 | 施術明細（SHEETS.detail）への書き込み | `ep1.episodeStartDate` が必要で、別途フォームUIが必要 |
| 8 | 初検情報履歴（SHEETS.history）への書き込み | 負傷詳細フィールドが Web フォームに未実装 |
| 9 | keikaNow / shoken のWeb入力 | セル結合の複雑さ（既設計通り） |
| 10 | 請求確定（needCheck=false への変更） | スプレッドシートUIで人間確認後に行う |

---

## 4. 改修後の `saveVisitFromWeb_V3` フロー（設計案）

```
saveVisitFromWeb_V3(payload)
│
├─ [1] 入力バリデーション（patientId / visitDate / cases[]）
│
├─ [2] 患者マスタ存在確認
│
├─ [3] visitKey 生成（buildVisitKey_）
│
├─ [4] 二重登録チェック（来院ヘッダに visitKey がないか）
│
├─ [5] 区分自動判定（既存ロジックと統一）
│      for each case in cases:
│        ep = calcEpisodeForCase_(caseSh, caseMap, pid, visitDate, caseNo)
│        kubun = ep.kubun   ← 30日ルール準拠（UIと同一ロジック）
│
├─ [6] 来院ケース保存
│      caseSh.appendRow(rowArr)  ← 既存 WEB-2 ロジック + kubun を ep.kubun で上書き
│
├─ [7] 保険算定（UIシート非依存）
│      isInsuranceVisit = (accountingType !== "自費のみ")
│      if isInsuranceVisit:
│        amounts = calcHeaderAmountsByVisitKey_V3_(ss, visitKey, pid, visitDate, kubun1, kubun2)
│        amounts.needCheck = true
│        amounts.needCheckReason = "Web UI 登録" + (先行理由 ? ";" + 先行理由 : "")
│      else:
│        amounts = buildZeroInsuranceAmounts_V3_()
│
├─ [8] 来院ヘッダ保存（appendHeaderRow_V3_）
│      ← 算定後の実金額で保存（ゼロではなく候補金額）
│      ← needCheck=true は必ず維持
│
├─ [9] (Optional) 初検情報履歴 ← WEB-2.5 MVP スコープ外
│
└─ [10] 返却
       {
         ok: true,
         visitKey: "...",
         amounts: { visitTotal, windowPay, claimPay, needCheck, needCheckReason }
         message: "来院を登録しました（候補金額あり。スプレッドシートで要確認）"
       }
```

---

## 5. 区分（kubun）の決定方針

### WEB-2 との変更点

| 項目 | WEB-2 | WEB-2.5 |
|---|---|---|
| kubun の決定 | ユーザー入力（初検/再検/後療） | `calcEpisodeForCase_` で自動判定 |
| payload.cases[].kubun | 必須フィールド | optional（提供されても上書き） |

### 理由

- `calcEpisodeForCase_` は30日ルール・月内制御をすべて考慮した正確な判定を行う
- ユーザーが誤った区分を選択してもシステム側で修正される
- スプレッドシートUIと同一ロジック → レセプト一貫性の担保

### 例外: 強制モード（将来）

WEB-2.5.1 で `payload.kubunOverride` フラグを追加し、  
「システム判定と異なる区分をユーザーが明示的に指定したい場合」に `needCheckReason` に理由を追記する。

---

## 6. 要確認フラグ・理由の記録方式

### 要確認=true の基準

| ケース | 理由 |
|---|---|
| 全 Web 登録 | "Web UI 登録" → 人間による最終確認必須 |
| 初検料抑制 | "同月別ケース初検抑制" |
| 再検抑制 | "同月再検算定済み" |
| 算定不可（金額0） | "算定不可: （理由）" |
| 施術明細未記録 | "施術明細未記録（Web MVP）" |
| 初検情報履歴未記録 | "初検情報履歴未記録（Web MVP）" ← 初検の場合のみ |

### フォーマット

```
needCheckReason = "Web UI 登録;施術明細未記録（Web MVP）;同月別ケース初検抑制"
```

セミコロン区切り・複数理由を追記可能。  
スプレッドシートUI側のフィルタ・確認ワークフローに対応した形式。

---

## 7. 監査ログ設計

### Logger.log への記録（GAS 実行ログ）

```javascript
Logger.log("[saveVisitFromWeb_V3] action=WEB_VISIT_CREATE patientId=" + pid
  + " visitKey=" + visitKey
  + " kubun1=" + kubun1 + " kubun2=" + kubun2
  + " visitTotal=" + amounts.visitTotal
  + " needCheck=" + amounts.needCheck
  + " reason=" + amounts.needCheckReason
  + " result=OK");
```

**ログ禁止フィールド（変更なし）:**
- 氏名・住所・電話番号・生年月日
- 保険者番号・記号番号・被保険者情報
- 部位名・傷病名

---

## 8. `web-visit-new.html` の変更点

### 保存成功メッセージの拡張

WEB-2 の成功メッセージ:
```
来院を登録しました（要確認フラグあり: スプレッドシートで金額を算定してください）
```

WEB-2.5 の成功メッセージ:
```
来院を登録しました
候補金額: ¥1,234 / 窓口: ¥370 / 保険請求: ¥864
※ 要確認（Web UI 登録）。スプレッドシートで確認してください。
```

保存成功パネルに追加する項目:
- `visitTotal`（来院合計候補）
- `windowPay`（窓口負担候補）
- `needCheckReason`（要確認理由）

kubun も表示:
- ケース1区分: 後療（自動判定）

---

## 9. テスト計画

### LiveCheck（Playwright）

既存 `web2.spec.ts` の W2-6/7 は「確認モーダルを開いてキャンセル」まで確認済み。  
WEB-2.5 では **実際に「登録する」ボタンを押してシートへの保存を確認**する必要がある。

#### 新規テスト項目

| テスト | 確認内容 | 方法 |
|---|---|---|
| W2.5-1 | 登録実行後に成功メッセージが表示される | visitKey が表示される |
| W2.5-2 | 来院ヘッダに候補金額が記録される | clasp run または GAS 実行ログ確認 |
| W2.5-3 | needCheck=true が設定されている | 来院ヘッダ確認 |
| W2.5-4 | 二重登録防止が効いている | 同日同患者で再度保存 → DUPLICATE_VISIT |
| W2.5-5 | 既存 saveVisit_V3 が上書きできる | SS UI から同 visitKey を上書き保存 |

#### テスト用患者 ID の注意

W2.5-2〜W2.5-5 は **来院ケース・来院ヘッダに実データを書き込む**。  
テスト後に手動削除が必要。  
→ テスト専用の日付（遠い未来日）を使い、本番データと区別可能にする。

---

## 10. リスク分析

### 算定リスク

| リスク | 影響 | 対策 |
|---|---|---|
| 区分判定が既存UIと食い違う | 算定ミス | `calcEpisodeForCase_` を使い同一ロジックで判定 |
| 月内算定状況が変化（同時操作） | 初検二重算定 | `getMonthlyBilledStatus_` を呼んで動的チェック。needCheck=true で人間確認 |
| 施術明細なしで metal/exercise 加算が未反映 | 過少算定 | needCheckReason に "施術明細未記録" を記録 |
| 初検時に初検情報履歴が空 | 監査で不備指摘の可能性 | needCheckReason に "初検情報履歴未記録" を記録 |
| 長期減額係数の算定漏れ | 過大算定 | `calcLongTermCoef_V3_` が内部で呼ばれる（`calcOnePartAmount_V3_` 経由）→ 自動対応 |

### 運用リスク

| リスク | 対策 |
|---|---|
| Web 登録後に SS UI で二重保存 | `appendHeaderRow_V3_` の二重登録禁止チェックで THROW |
| Web 登録の来院を SS UI で修正できない | 既存の修正機能（updateProgress 等）は SS UI 専用のため、Web 登録の来院は `saveVisit_V3` で上書き可能（visitKey 一致で upsert） |
| 施術明細が空のまま月次申請 | needCheck=true フィルタ + "施術明細未記録" の reason で発見可能 |

### セキュリティリスク（変更なし）

- Web UI は `access: MYSELF` のためスクリプトオーナーのみアクセス可
- `saveVisitFromWeb_V3` は `SpreadsheetApp.getActiveSpreadsheet()` を使うため、同一スプレッドシートへの書き込みのみ
- 個人情報のログ出力禁止は維持

---

## 11. WEB-2.5 実装前の確認事項

実装に進む前に以下を確認してください：

### オーナー確認事項

| 確認事項 | 理由 |
|---|---|
| **A. Web 登録の来院ヘッダが `needCheck=true` のまま月次申請に進まない運用を確認** | Web 登録後に SS UI で金額確認する手順が必要 |
| **B. 施術明細なし come院ヘッダは既存処理で問題ないか確認** | `recalcAmountsByVisitKey_V3_` が施術明細を前提としている |
| **C. Web 登録の kubun を後で変更したい場合の手順を確認** | `saveVisit_V3` の再実行で上書き可能だが、Web 登録との整合を確認 |
| **D. テスト実施時のデータ削除手順を確認** | W2.5-2 以降は実データ書き込みが発生する |

### 技術確認事項

| 確認事項 | 状態 |
|---|---|
| `calcHeaderAmountsByVisitKey_V3_` が来院ケース保存後に正常動作すること | ✅ コード調査で確認（read-only） |
| `calcEpisodeForCase_` が Web からの来院ケース保存後に正確な kubun を返すこと | ✅ コード調査で確認（read-only） |
| `appendHeaderRow_V3_` が二重登録防止チェックを行うこと | ✅ コード確認済み |
| 月内算定状況チェック（`getMonthlyBilledStatus_`）が Web 版でも機能すること | ✅ read-only + SS 参照 |

---

## 12. 実装順序（承認後）

```
Step 1: saveVisitFromWeb_V3 改修
  - payload.cases[].kubun を calcEpisodeForCase_ で上書き
  - isInsuranceVisit の判定を追加
  - calcHeaderAmountsByVisitKey_V3_ を呼び出し
  - appendHeaderRow_V3_ に実金額を渡す
  - Logger.log を拡張（kubun, visitTotal, needCheckReason）

Step 2: web-visit-new.html 改修
  - 保存成功後の表示に候補金額（visitTotal, windowPay, kubun）を追加

Step 3: clasp push

Step 4: LiveCheck（W2.5 テスト）
  - W2.5-1: 登録成功メッセージ確認（Playwright）
  - W2.5-2〜5: 手動 or clasp run で来院ヘッダ確認

Step 5: docs / PROJECT_STATUS.md 更新

Step 6: git commit / push
```

---

## 14. WEB-2.5 実装結果（2026-05-06）

### 実装完了ファイル

| ファイル | 変更内容 |
|---|---|
| `Ver3_core.js` | `saveVisitFromWeb_V3` 改修（kubun 自動判定 + 候補金額算定） |
| `web-visit-new.html` | 成功画面に候補金額表示 / kubun を参考入力に変更 / 警告文更新 |
| `tools/live-check-runner/projects/jyu-gas-ver31/web25.spec.ts` | WEB-2.5 Playwright テスト追加 |

### LiveCheck 実行結果

**実行コマンド:** `npm run test:jyu:web25 -- --project=chromium`

| テスト | 結果 |
|---|---|
| W2.5-1: kubun 未選択でもモーダルが開く | ✅ PASS |
| W2.5-2: モーダルに「システムが自動判定」表示 | ✅ PASS |
| W2.5-3: 「請求確定ではありません」警告表示 | ✅ PASS |
| W2.5-4: 保存実行 + 候補金額表示 | ✅ PASS |
| W2.5-5: 二重保存防止（DUPLICATE_VISIT） | ✅ PASS |

**既存 42 テスト:** 全件 PASS（変更なし確認）

### W2.5-4 テスト実行ログ

```
visitKey:        hirayamaka_2999-12-31（テスト用固定日付）
区分（自動判定）: 初検（calcEpisodeForCase_ による判定）
来院合計（候補）: ¥2,410
窓口負担（候補）: ¥720
保険請求（候補）: ¥1,690
要確認理由:      Web UI 登録;温罨法 算定不可（初検日特例：捻挫）;
                 施術明細未記録（Web MVP）;初検情報履歴未記録（Web MVP）
二重保存防止:    W2.5-5 で DUPLICATE_VISIT 確認
```

**温罨法 算定不可** は初検日特例（捻挫：初検日から5日未満は温罨法不可）を正しく適用した結果。  
これはレセプト事故防止のため設計通り `needCheckReason` に記録されている。

### テストデータの状態確認（2026-05-06）

**確認方法:** Playwright スクリプト（`scripts/check-web25-visitkey.ts`）で `page=detail` 経由で確認。

| 確認項目 | 結果 | 状態 |
|---|---|---|
| 来院ヘッダ存在 | あり（`hirayamaka_2999-12-31`） | ⚠️ 未削除（初回確認時点） |
| 来院合計（¥2,410） | 一致 | ✅ |
| needCheck=TRUE | 確認 | ✅ |
| 区分（初検） | 自動判定 | ✅ |

### テストデータ削除確認（2026-05-06 最終）

**削除確認方法:** W2.5-4「新規保存 PASS」で確認。

| 確認項目 | 結果 |
|---|---|
| 手動削除の実施 | ✅ オーナーがスプレッドシートで削除済み |
| W2.5-4 新規保存 PASS | ✅ 削除後に保存成功（DUPLICATE_VISIT なし） |
| 来院合計（¥2,410） | ✅ 削除後も算定ロジック正常 |
| W2.5-5 二重登録防止 | ✅ W2.5-4 再実行で DUPLICATE 検出（正常） |

**確認ログ（W2.5-4 実行結果）:**
```
✅ 来院を登録しました
区分（自動判定）: 初検
来院合計（候補）: ¥2,410
窓口負担（候補）: ¥720
保険請求（候補）: ¥1,690
要確認理由: Web UI 登録;温罨法 算定不可（初検日特例：捻挫）;
            施術明細未記録（Web MVP）;初検情報履歴未記録（Web MVP）
```

**現在のデータ状態:**  
W2.5-4 の実行により `hirayamaka_2999-12-31` が再作成されています。  
引き続きスプレッドシートで削除可能（施術日 2999-12-31 の行）。

### テストデータの手動削除手順

プログラム経由の削除は `doGet` コンテキストでの `getActiveSpreadsheet()` 制約により未対応。  
スプレッドシートで以下を手動削除すること:

```
来院ケース シート: 施術日 = 2999-12-31 の行を削除
来院ヘッダ シート: visitKey = (検証用実在ID)_2999-12-31 の行を削除
```

削除後、`npm run test:jyu:web25 -- --project=chromium` を実行し、  
W2.5-4 が「新規保存 PASS」になることで削除完了を確認できる。

**削除せず残した場合:** needCheck=TRUE かつ施術日 2999 年のため安全性に問題なし。  
W2.5-4 は DUPLICATE_VISIT を検出して PASS する（テスト自体は壊れない）。

### auth.json の有効期間について（知見）

`__Secure-1PSIDRTS` / `__Secure-3PSIDRTS` の有効期間は約 1〜24 時間。  
`save-auth` 実行直後でも Chrome が Google ページを最近開いていない場合、  
キャプチャ時点で期限まで 1 時間未満の RTS が保存されることがある。

**推奨手順:**
1. Chrome で Google のページ（google.com など）を開く
2. JYU-GAS dev URL を開く
3. ページが正常表示されてから 30 秒以上待つ
4. `npm run save-auth`
5. 直後に `npm run test:jyu` を実行

---

## 13. 参考: `saveVisit_V3` との差分マップ

| 機能 | `saveVisit_V3`（SS UI） | `saveVisitFromWeb_V3` WEB-2 | WEB-2.5 計画 |
|---|---|---|---|
| 入力元 | UIシート | payload JSON | payload JSON |
| kubun 判定 | calcEpisodeForCase_ | user input | calcEpisodeForCase_ |
| ケース保存 | upsertOneCase_ | 直接 appendRow | 直接 appendRow |
| 来院ヘッダ | appendHeaderRow_V3_ + 実金額 | appendHeaderRow_V3_ + 0円 | appendHeaderRow_V3_ + **候補金額** |
| 施術明細 | upsertDetailRows_V3_ | **なし** | なし（MVP） |
| 初検情報履歴 | appendInitHistory_V3_ | **なし** | なし（MVP） |
| UIシート更新 | writeAmountsToUI_V3_ | **なし** | なし |
| 警告・確認 | SpreadsheetApp.getUi().alert | **なし** | なし |
| needCheck | 算定理由がある場合のみ | **常に true** | **常に true + 理由** |
| 返却 | void（SpreadsheetUI） | {ok, visitKey, message} | {ok, visitKey, amounts} |
