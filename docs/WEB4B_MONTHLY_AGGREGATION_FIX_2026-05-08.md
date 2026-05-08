# WEB-4B: 月次申請集計0円バグ修正記録

作成日: 2026-05-08  
ステータス: **修正完了 / clasp push 済 / commit & push 済 / 本番 deploy 未実施**

---

## 問題

月次申請一覧（`?page=monthlyClaims`）および月次申請詳細（`?page=monthlyClaimDetail`）で、
2026-04 の 9名全員の集計が来院数0・金額¥0・状態「来院なし」と表示されていた。

B案 Cloud Run Excel 申請書生成（`generateClaimApplicationBFromWeb_V3`）は正常動作しており、
Web 表示集計ロジックだけが誤っていた。

---

## 根本原因

`getMonthlyClaimList_V3` と `getMonthlyClaimDetail_V3`（Ver3_core.js）が、
来院ヘッダシートの列マップ取得に **`buildHeaderColMap_`** を使用していた。

| 関数 | 返り値 | 用途 |
|---|---|---|
| `buildHeaderColMap_` | **1始まり** 列番号（GAS `getRange` 用） | Range操作向け |
| `V3TR_buildHeaderMap_` | **0始まり** インデックス（配列アクセス用） | `getValues()` 配列向け |

`buildHeaderColMap_` が返す1始まりの値を、`getValues()` の0始まり配列インデックスとして使った結果、
常に1列ずれた値（または `undefined`）を参照した。

```javascript
// ❌ 修正前 (バグ)
var headMap = buildHeaderColMap_(headSh);     // 1始まり
var hDtC    = headMap["施術日"];               // e.g., 2 (列Bが1始まりで2番)
var hdt     = hVals[hr][hDtC];               // hVals[hr][2] → 列C を参照 (0始まり) WRONG
// → hdt は Date でない → V3TR_inRange_ が false → 全行スキップ → 集計0
```

```javascript
// ✅ 修正後
var headMap = V3TR_buildHeaderMap_(headSh);   // 0始まり
var hDtC    = headMap["施術日"];               // e.g., 1 (列Bが0始まりで1番)
var hdt     = hVals[hr][hDtC];               // hVals[hr][1] → 列B を参照 CORRECT
// → hdt は Date → V3TR_inRange_ が正常動作 → 集計正常
```

`V3TR_findPatientsForMonth_`（B案で使う患者抽出 = 正常動作）は最初から
`V3TR_buildHeaderMap_`（0始まり）を使っていたため正しく動いていた。
この非一貫性が原因で「患者は9名見つかるが集計は全員0」という現象になっていた。

---

## 影響範囲

| 対象 | 変更 | 影響 |
|---|---|---|
| `getMonthlyClaimList_V3` | `buildHeaderColMap_` → `V3TR_buildHeaderMap_` | 月次一覧の集計が正常に |
| `getMonthlyClaimDetail_V3` | 同上 | 月次詳細の集計が正常に |
| B案申請書生成ルート | **変更なし** | 影響なし |
| `V3TR_menuGenerateApplication_B` | **変更なし** | 影響なし |
| `generateClaimApplicationBFromWeb_V3` | **変更なし** | 影響なし |
| その他の `buildHeaderColMap_` 使用箇所 | **変更なし** | 正しくRange操作で使用中 |

---

## 修正内容

**ファイル:** `Ver3_core.js`

**変更箇所 1:** `getMonthlyClaimList_V3` 内（来院ヘッダ集計部）

```javascript
// Before
var headMap = buildHeaderColMap_(headSh);
// After
// V3TR_buildHeaderMap_ を使用（0始まりインデックスで配列アクセスと一致）
var headMap = V3TR_buildHeaderMap_(headSh);
```

**変更箇所 2:** `getMonthlyClaimDetail_V3` 内（詳細集計部）

```javascript
// Before
var headMap = buildHeaderColMap_(headSh);
// After
// V3TR_buildHeaderMap_ を使用（0始まりインデックスで配列アクセスと一致）
var headMap = V3TR_buildHeaderMap_(headSh);
```

---

## 確認結果

### 実測値（2026-04 / hirayamaka）

| 項目 | 修正前 | 修正後 |
|---|---|---|
| 来院数 | 0 | **2** |
| 来院合計 | ¥0 | **¥4,368** |
| 窓口負担 | ¥0 | **¥1,310** |
| 保険請求 | ¥0 | **¥3,058** |
| 状態 | 来院なし | 要確認あり（1件）|

### LiveCheck（WEB-4B）

| テスト | 内容 | 結果 |
|---|---|---|
| W4B-1 | monthlyClaims 2026-04 で患者が存在する | ✅ PASS |
| W4B-2 | monthlyClaimDetail 来院数が 0 でない | ✅ PASS |
| W4B-3 | monthlyClaimDetail 保険請求が 0 でない | ✅ PASS |
| W4B-4 | 来院明細行が存在する | ✅ PASS |
| W4B-5 | B案申請書生成ボタンが引き続き存在する | ✅ PASS |

### 回帰テスト

| スイート | 結果 |
|---|---|
| web3 | **8 PASS** |
| web4 (WEB-4A) | **5 PASS** |
| web4b (WEB-4B) | **5 PASS** |

---

## clasp push

実施済み（2026-05-08 11:03:17）

---

## 本番 deploy

**未実施（本番 deploy 禁止のため）**

---

## Dashboard

**対象外**

---

## 残確認事項

- [ ] `/exec?page=monthlyClaims` で本番環境でも集計が正しく表示されるか（本番 deploy 後に確認）
- [ ] 全 9 患者の一覧集計も正しく表示されるか（dev URL で手動確認可）
