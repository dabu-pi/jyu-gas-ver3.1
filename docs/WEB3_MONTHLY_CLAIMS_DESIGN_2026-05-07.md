# WEB-3 月次申請フロー — 設計記録

作成日: 2026-05-07  
対象フェーズ: Phase WEB-3  
ステータス: WEB-3.1 / WEB-3.2 / WEB-3.3(転記) 実装完了・LiveCheck 8 PASS

---

## 実装概要

Web UI から月次申請対象者一覧の確認・プレビュー・転記データ生成を行えるようにした。

### 実装範囲

| フェーズ | 内容 | 状態 |
|---|---|---|
| WEB-3.1 | 月次申請対象者一覧 | ✅ 完了 |
| WEB-3.2 | 申請対象者詳細・プレビュー | ✅ 完了 |
| WEB-3.3 | 申請書転記データ生成（Web UI から呼び出し） | ✅ 完了 |
| WEB-3.4 | 申請書 PDF 生成・印刷 | ⏸ 未着手（スプレッドシートUIで実施） |

---

## 新規 GAS 関数

| 関数名 | 役割 |
|---|---|
| `getMonthlyClaimList_V3(ym)` | 指定月に保険来院がある患者一覧と集計を返す |
| `getMonthlyClaimDetail_V3(patientId, ym)` | 患者×月の来院詳細一覧（読み取り専用） |
| `buildMonthlyTransferData_V3(patientId, ym)` | 既存 `V3TR_buildTransferDataForMonth_` を Web から呼び出せるようラップ |

### 既存ロジックの再利用

| 既存関数 | 再利用箇所 |
|---|---|
| `V3TR_findPatientsForMonth_(ss, ym)` | `getMonthlyClaimList_V3` — 対象月に来院がある患者 ID 一覧取得 |
| `V3TR_parseYM_(ym)` | 月範囲 (start/end) の計算 |
| `V3TR_inRange_(dt, start, end)` | 来院ヘッダの日付フィルタリング |
| `V3TR_buildTransferDataForMonth_(ss, pid, ym)` | `buildMonthlyTransferData_V3` の実装 |
| `buildHeaderColMap_(sh)` | ヘッダ列マップ構築 |
| `HEADER_COLS.*` | 来院ヘッダ列名定数 |

重複実装なし。既存の月次集計・算定ロジックを 100% 再利用。

---

## 新規 HTML ページ

| ファイル | page= | 内容 |
|---|---|---|
| `web-monthly-claims.html` | `monthlyClaims` | 月次申請対象者一覧（年月選択 → テーブル表示） |
| `web-monthly-claim-detail.html` | `monthlyClaimDetail` | 患者×月の詳細・プレビュー・転記データ生成ボタン |

---

## doGet ルーティング追加

| page= | HTML | 追加パラメータ |
|---|---|---|
| `monthlyClaims` | `web-monthly-claims.html` | なし |
| `monthlyClaimDetail` | `web-monthly-claim-detail.html` | patientId, ym |

---

## 設計上の重要点

### 制度遵守

| 原則 | 実装 |
|---|---|
| 算定事実主義 | 来院ヘッダの確定値（visitTotal / claimPay）のみ表示。推計・概算なし |
| needCheck 維持 | Web 登録来院は常に needCheck=true。一覧画面で「要確認」バッジを表示 |
| 請求確定は人間が行う | Web UI では「詳細確認」と「転記データ生成」まで。請求確定は Sheets UI |
| 自費のみ来院除外 | `accountingType === "自費のみ"` の行は申請対象外として一覧から除外 |

### データの流れ

```
来院ヘッダ → getMonthlyClaimList_V3
          → getMonthlyClaimDetail_V3（詳細）
          → buildMonthlyTransferData_V3
               └→ V3TR_buildTransferDataForMonth_
                    └→ 申請書_転記データシート（upsert）
                    └→ スプレッドシートUIで申請書出力
```

### 監査ログ

- GAS 側 Logger.log で各操作（一覧取得・詳細取得・転記データ生成）を記録
- 転記データ生成は `buildMonthlyTransferData_V3` でログを出力
- 監査ログ構造は将来フェーズで整備（現在は Logger のみ）

---

## doGet 変更

既存ルートへの影響なし。追加のみ。

```
page=search        → patientSearch.html   （変更なし）
page=selfpay       → selfPayWeb.html       （変更なし）
page=home          → web-home.html         （カード更新のみ）
page=detail        → web-patient-detail.html（変更なし）
page=visitNew      → web-visit-new.html   （変更なし）
page=monthlyClaims → web-monthly-claims.html     ← WEB-3.1 追加
page=monthlyClaimDetail → web-monthly-claim-detail.html ← WEB-3.2 追加
```

---

## web-home.html 更新

| カード | 変更前 | 変更後 |
|---|---|---|
| 来院記録 | disabled（Phase WEB-2 バッジ） | active リンク → ?page=visitNew |
| 申請書 | disabled（Phase WEB-3 バッジ） | 削除し「月次申請」セクションに統合 |
| 月次申請 | なし | 新規追加（active リンク → ?page=monthlyClaims） |

---

## 申請書 PDF 生成について（WEB-3.4）

申請書 PDF 生成は既存の `V3TR_menuGenerateApplication_B` / `V3TR_writeToApplication_` が担う。
これらは Spreadsheet UI 専用（SpreadsheetApp.getUi().prompt 使用）のため、Web App からは呼び出せない。

**現時点の運用フロー:**
1. Web UI → 申請書転記データ生成（`buildMonthlyTransferData_V3`）
2. スプレッドシート → 「申請書を出力」メニュー → PDF 生成
3. 将来: Web App 専用の申請書出力フローを追加（WEB-3.4）

---

## LiveCheck 結果（2026-05-07）

```
npm run test:jyu:web3
W3-1: PASS — 月次申請カードが home に表示される
W3-2: PASS — monthlyClaims ページ表示
W3-3: PASS — 年月入力欄（デフォルト: 2026-05）
W3-4: PASS — 一覧を取得ボタン存在
W3-5: PASS — monthlyClaimDetail パラメータなしでもクラッシュしない
W3-6: PASS — patientSearch 既存導線正常
W3-7: PASS — active カード数 3（患者検索 + 来院記録 + 月次申請）
W3-8: PASS — iframe 入れ子増殖なし（iframe 数 1）

8 PASS / 0 FAIL / 0 SKIP
```

回帰テスト:
```
smoke:  28 PASS
web25:   5 PASS
web251:  3 PASS / 1 SKIP（テストデータ 2998-12-31 残存による正常 SKIP）
web3:    8 PASS
合計:   44 PASS / 1 SKIP / 0 FAIL
```

---

## テストデータ残存（要手動削除）

| visitKey | 由来 | 削除対象シート |
|---|---|---|
| hirayamaka_2998-12-31 | WEB-2.5.1 テスト (W2.5.1-1) | 来院ケース・来院ヘッダ・施術明細 |
| hirayamaka_2999-12-31 | WEB-2.5 テスト (W2.5-4) | 来院ケース・来院ヘッダ |

いずれも未来日付（2998年・2999年）のため月次申請の実運用には影響なし。  
スプレッドシートで施術日列を確認して削除すること。
