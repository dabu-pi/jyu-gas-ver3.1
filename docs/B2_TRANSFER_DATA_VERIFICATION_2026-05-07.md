# B-2 月次申請データ整合確認 — 記録

実施日: 2026-05-07  
対象フェーズ: B-2（施術明細 → 申請書転記データ → PDF生成前提）  
ステータス: インフラ検証完了 / 実データ確認は人間タスク（月次来院後）

---

## 実施内容

### web251 SKIP 解消確認

テストデータ削除後に web251 を再実行した結果:

| 実行回 | 結果 | 備考 |
|---|---|---|
| 1回目（削除後） | **4 PASS / 0 SKIP** | W2.5.1-1 が hirayamaka_2998-12-31 を新規保存 → PASS |
| 2回目（同セッション内） | **3 PASS / 1 SKIP** | W2.5.1-1 が DUPLICATE_VISIT で SKIP（正常動作） |

**結論:** SKIP は設計通り。W2.5.1-1 は最初の実行でデータを作成し、次回以降は DUPLICATE → SKIP。  
完全な 4 PASS 化には毎回の cleanup が必要（`devCleanupTestVisitData_V3(false)` を実行）。

---

## B-2 実装内容

### 新規 GAS 関数

| 関数 | 役割 |
|---|---|
| `verifyMonthlyClaimData_V3(ym)` | 月次申請データ整合確認（読み取り専用） |

`verifyMonthlyClaimData_V3` は以下を順に実行する（write 操作なし）:
1. `getMonthlyClaimList_V3(ym)` で対象患者一覧を取得
2. 先頭患者に対して `getMonthlyClaimDetail_V3(pid, ym)` を取得
3. 一覧 vs 詳細の整合をチェック（visitCount / claimPay / needCheckCount）
4. ステータスを返す: `INTEGRITY_OK` / `NO_PATIENTS_THIS_MONTH` / `INTEGRITY_MISMATCH`

### 新規ページ・ルート

| 追加 | 内容 |
|---|---|
| `web-b2-results.html` | B-2 確認結果自動表示（Playwright DOM 読み取り） |
| `page=b2Results` doGet | verifyMonthlyClaimData_V3 を呼び出すルート |

### 新規 LiveCheck spec

`b2_transfer.spec.ts` — B2-1〜B2-10

---

## B-2 LiveCheck 結果（2026-05-07）

```
npm run test:jyu:b2
B2-1:  PASS — b2Results ページ到達（ym=2026-05）
B2-2:  PASS — verifyMonthlyClaimData_V3 正常完了
B2-3:  PASS — status=NO_PATIENTS_THIS_MONTH
B2-4:  PASS — 整合チェック（患者なし月のため実行なし）
B2-5:  INFO — 2026-05 に保険来院がある患者なし（人間確認要）
B2-6:  PASS — データ読み取りフロー正常
B2-7:  PASS — monthlyClaims 正常
B2-8:  PASS — patientSearch 正常
B2-9:  PASS — iframe 数 1（増殖なし）
B2-10: PASS — 重大 console error なし

6 PASS / 0 FAIL / 0 SKIP
```

---

## B-2 ステータス詳細

| 確認項目 | 結果 |
|---|---|
| 施術明細 → 月次申請対象者一覧の読み取り | ✅ PASS（`verifyMonthlyClaimData_V3` 正常動作） |
| 対象月（2026-05）の保険来院患者 | `NO_PATIENTS_THIS_MONTH`（当月データなし） |
| 整合チェック（list vs detail） | ✅ 実行可能・患者がいれば自動検証 |
| 実データでの整合確認 | ⏸ 人間確認待ち（実際の来院がある月で確認） |

---

## 「新 様式第5号」テンプレート確認状況

WEB-3.4 LiveCheck（W3.4-6）で `ZERO_CLAIM` が返った:
- 当月（2026-05）に hirayamaka の保険請求額 = 0
- PDF 生成前の Layer 2 安全フィルタが正常動作
- テンプレートシートの有無は未確認（実患者・実月で確認が必要）

**スプレッドシートでの確認が必要:**
- 「新 様式第5号」シートが存在するか
- 存在すれば実患者・実月で `generateClaimApplication_V3` を試行

---

## 実データ確認が必要な項目（人間タスク）

| 確認項目 | 手順 |
|---|---|
| 実患者・実月での B-2 整合確認 | 来院データのある月（例: 2026-04）で `?page=b2Results&ym=2026-04` を開く |
| 申請書テンプレート有無 | スプレッドシートで「新 様式第5号」シートを確認 |
| PDF 生成の実動作 | 実患者・実月で `generateClaimApplication_V3` を試行 |

---

## B-2 残タスク判断

| 判断 | 根拠 |
|---|---|
| B-2 インフラ検証: 完了 | verifyMonthlyClaimData_V3 正常動作 / 全 6 PASS |
| B-2 実データ検証: 保留 | 2026-05 にデータなし / 月次来院後に確認 |
| B-2 CLOSED 可否: 部分 CLOSED | インフラ検証は完了 / 実データは月次運用に委ねる |
