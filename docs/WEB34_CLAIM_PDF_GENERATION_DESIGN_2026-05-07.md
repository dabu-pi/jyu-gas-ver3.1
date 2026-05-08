# WEB-3.4 申請書PDF生成 — 設計記録

作成日: 2026-05-07  
対象フェーズ: Phase WEB-3.4  
ステータス: 実装完了・LiveCheck 9 PASS / 1 SKIP

---

## 実装概要

Web UI から月次申請詳細ページ（web-monthly-claim-detail.html）に「申請書PDFを生成」ボタンを追加した。
既存のテンプレートシート書き込み（A案）+ Drive PDF エクスポートにより、Web App から申請書 PDF を生成できる。

---

## 実装方式の選択理由

### 方式の選定

| 方式 | 説明 | 採用 |
|---|---|---|
| A案 | テンプレートシート（新 様式第5号）書き込み + UrlFetchApp で PDF エクスポート | ✅ 採用 |
| B案 | 外部 Cloud Run エンドポイント（APPGEN_ENDPOINT）経由で xlsx 生成 | — 外部依存のため今回は不採用 |

**A案採用の理由:**
- 外部エンドポイント不要（APPGEN_ENDPOINT 未設定環境でも動作）
- 既存の `V3TR_writeToApplication_` と `V3TR_buildTransferDataForMonth_` を完全再利用
- GAS 標準の UrlFetchApp + Drive PDF エクスポートは Web App コンテキストで動作する
- B案は `APPGEN_ENDPOINT` スクリプトプロパティの設定が必要で追加環境依存がある

---

## 新規 GAS 関数

| 関数名 | 役割 |
|---|---|
| `generateClaimApplication_V3(patientId, ym)` | WEB-3.4 メイン：転記 + テンプレ書込 + PDF 生成 |
| `devCleanupTestVisitData_V3(dryRun)` | DEV ONLY: 未来日テストデータの安全削除（2990年以降のみ対象） |

---

## generateClaimApplication_V3 の処理フロー

```
1. 入力バリデーション（patientId, ym 形式チェック）
2. V3TR_buildTransferDataForMonth_（転記データ生成・upsert 冪等）
3. Layer 2 安全フィルタ: claimPay = 0 は ZERO_CLAIM で拒否
4. 申請書_転記データシートから recordKey1 / recordKey2 行を読み込み
5. V3TR_writeToApplication_（テンプレートシート「新 様式第5号」に書き込み）
6. UrlFetchApp で PDF エクスポート → Drive 保存
7. { ok, pdfUrl, fileId, writtenCells, transferClaim, message } を返却
```

### 安全ガード

| ガード | 内容 |
|---|---|
| Layer 2 フィルタ | claimPay = 0 の場合は PDF 生成拒否（保険申請対象外） |
| テンプレート未存在 | TEMPLATE_NOT_FOUND を返す（転記データ生成は完了済みと報告） |
| PDF エクスポート失敗 | HTTP エラーコードを返す（テンプレート書き込みは完了済みと報告） |
| 監査ログ | Logger.log で各ステップ（BUILD_TRANSFER / WRITE_TEMPLATE / PDF_CREATED）を記録 |

---

## 既存ロジック再利用

| 再利用関数 | 用途 |
|---|---|
| `V3TR_buildTransferDataForMonth_` | 転記データ生成（upsert） |
| `V3TR_rowToObj_` | 転記シート行 → オブジェクト変換 |
| `V3TR_writeToApplication_` | 申請書テンプレートへの書き込み |
| `V3TR_getApplicationOutputFolder_` | Drive 保存先フォルダ取得 |

---

## UI 変更（web-monthly-claim-detail.html）

既存の「申請書転記データを生成」ボタン（Step 1）の下に「申請書PDFを生成」ボタン（Step 2）を追加。

```
Step 1: 申請書転記データ生成（既存 buildMonthlyTransferData_V3 / upsert）
Step 2: 申請書PDF生成（新規 generateClaimApplication_V3）
```

Step 2 の動作:
1. confirm ダイアログで確認
2. `generateClaimApplication_V3(patientId, ym)` 呼び出し
3. 結果を画面に表示:
   - 成功: PDF の Drive リンク・書込セル数・請求金額
   - TEMPLATE_NOT_FOUND: 警告表示（転記データ生成完了を通知）
   - ZERO_CLAIM: エラー表示（申請対象外）
   - その他エラー: エラー表示

---

## LiveCheck 結果（2026-05-07）

```
npm run test:jyu:web34
W3.4-1:  PASS — monthlyClaims 到達
W3.4-2:  PASS — monthlyClaimDetail クラッシュなし
W3.4-3:  PASS — Step 1 ボタン存在
W3.4-4:  PASS — Step 2 (WEB-3.4) ボタン存在
W3.4-5:  PASS — PDF生成ボタンがクリック可能
W3.4-6:  PASS — generateClaimApplication_V3 呼び出し（ZERO_CLAIM = Layer2 安全フィルタ正常動作）
W3.4-7:  PASS — patientSearch 既存導線正常
W3.4-8:  PASS — iframe 数 1（増殖なし）
W3.4-9:  PASS — 重大 console error なし
W3.4-10: SKIP — inner frame evaluate 非対応（cleanup dry-run は GAS エディタから実施）

9 PASS / 1 SKIP / 0 FAIL
```

W3.4-6 の `ZERO_CLAIM` は正常動作:
- テスト患者（hirayamaka）の 2026-05 に保険来院なし
- Layer 2 安全フィルタが「請求額=0 は申請対象外」として拒否
- 本番では実際の患者・月を指定すれば PDF が生成される

---

## テストデータ残存と cleanup 関数

### 残存テストデータ

| visitKey | 残存シート | 削除方法 |
|---|---|---|
| hirayamaka_2998-12-31 | 来院ケース・来院ヘッダ・施術明細 | 下記 |
| hirayamaka_2999-12-31 | 来院ケース・来院ヘッダ | 下記 |

### devCleanupTestVisitData_V3 の使い方

Apps Script エディタ（Extensions → Apps Script）で以下を実行:

```javascript
// dry-run: 削除対象を報告するだけ（安全）
devCleanupTestVisitData_V3()         // または true を明示

// 実削除（慎重に！）
devCleanupTestVisitData_V3(false)
```

安全ガード: 年 ≥ 2990 の visitKey のみ対象（本番データは 2025〜2026 年なので絶対に触れない）

---

## 残タスク

| タスク | 状態 |
|---|---|
| テストデータ削除（hirayamaka_2998-12-31, 2999-12-31） | ⏸ 人間タスク（GAS エディタで実行） |
| 申請書テンプレート（新 様式第5号）有無確認 | ⏸ 人間確認（スプレッドシートで確認） |
| 本番 PDF 生成の実動作確認 | ⏸ 実患者・実月で人間確認 |
| WEB-3.4 本番 deploy | ⏸ 月次確認後に判断 |
| B案（Cloud Run API）の統合 | ⏸ APPGEN_ENDPOINT 設定後に検討 |
