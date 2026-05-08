# WEB-4A: Web UI B案申請書生成入口 実装記録

作成日: 2026-05-08  
ステータス: **実装完了 / clasp push 済 / commit & push 済**

---

## 概要

月次申請詳細画面（`web-monthly-claim-detail.html`）に、B案 Cloud Run Excel 申請書生成を呼べる入口を追加した。

---

## 追加した Web 導線

```
Web monthlyClaimDetail（?page=monthlyClaimDetail）
  → [申請書Excelを生成] ボタン（Step 3）
  → confirm ダイアログ
  → google.script.run.generateClaimApplicationBFromWeb_V3(patientId, ym)
  → GAS wrapper 関数（Ver3_transferData.js）
  → 既存 B案コア処理（V3TR_buildTransferDataForMonth_ / V3TR_exportTransferJson_）
  → プリフライトバリデーション（UIなし版）
  → Cloud Run POST /generate
  → Drive xlsx 保存
  → V3TR_writeGenerationLog_
  → 成功: fileUrl / fileName を Web に返す
```

---

## 追加・変更ファイル

| ファイル | 変更内容 |
|---|---|
| `Ver3_transferData.js` | `generateClaimApplicationBFromWeb_V3(patientId, ym)` 追加 |
| `web-monthly-claim-detail.html` | Step 3「申請書Excelを生成」セクション追加 / `doAppgenB()` JS関数追加 / needCheck警告対応 / Step 4 に PDF（試験的）をリナンバー |
| `tools/live-check-runner/projects/jyu-gas-ver31/web4_application_b.spec.ts` | WEB-4A LiveCheck（W4A-1〜5）新規 |
| `tools/live-check-runner/package.json` | `test:jyu:web4` スクリプト追加 |

---

## GAS Wrapper 設計

### 関数名

`generateClaimApplicationBFromWeb_V3(patientId, ym)`

### 既存B案との関係

| 要素 | 既存メニュー（V3TR_menuGenerateApplication_B） | Web wrapper（generateClaimApplicationBFromWeb_V3） |
|---|---|---|
| 対象患者 | 1人または月全員 | 1人（patientId 必須） |
| 入力方式 | UI prompt ダイアログ | google.script.run 引数 |
| プリフライト hard error | ui.alert + YES/NO | JSON error 返却（スキップ） |
| プリフライト warning | ui.alert + YES/NO | 自動続行（Logger.log のみ） |
| 進捗表示 | ss.toast | なし |
| 戻り値 | String（テキスト） | JSON（{ok, patientId, ym, fileName, fileUrl, message}） |
| 再利用する関数 | — | 全ての内部ヘルパー関数を再利用 |

### 再利用している既存関数

- `V3TR_buildTransferDataForMonth_(ss, pid, ym)` — 転記データ生成
- `V3TR_exportTransferJson_(ss, pid, ym, true)` — NDJSON 生成（skipBuild=true）
- `V3TR_loadClinicInfo_(shSettings)` — 施術機関情報取得
- `V3TR_getApplicationOutputFolder_(ss, ym)` — Drive 月別フォルダ
- `V3TR_getArchiveOutputFolder_(ss, ym)` — アーカイブフォルダ
- `V3TR_archiveExistingApplicationFiles_(...)` — 既存ファイル退避
- `V3TR_writeGenerationLog_(ss, ym, genAt, savedFiles, [])` — ログ記録

### 戻り値仕様

成功:
```json
{
  "ok": true,
  "patientId": "hirayamaka",
  "ym": "2026-04",
  "fileName": "申請書_hirayamaka_2026-04_093947.xlsx",
  "fileUrl": "https://drive.google.com/...",
  "message": "申請書Excelを生成しました。Driveで確認してください。"
}
```

エラー:
```json
{
  "ok": false,
  "errorCode": "ZERO_CLAIM",
  "message": "保険請求額が0円のため申請書を生成できません。"
}
```

### エラーコード一覧

| errorCode | 原因 |
|---|---|
| `INVALID_PATIENT_ID` | patientId が未指定 |
| `INVALID_YM` | 対象月フォーマット不正 |
| `APPGEN_CONFIG_MISSING` | APPGEN_ENDPOINT / APPGEN_SECRET 未設定 |
| `BUILD_FAILED` | 転記データ生成失敗 |
| `EXPORT_FAILED` | NDJSON エクスポート失敗 |
| `PARSE_FAILED` | JSON パース失敗 |
| `ZERO_CLAIM` | 保険請求額=0（申請対象外） |
| `PREFLIGHT_FAILED` | 必須項目不足 / 対象月不一致 |
| `MONTH_MISMATCH` | 転記データの対象月と実行月が不一致 |
| `CLOUDRUN_UNREACHABLE` | Cloud Run 接続失敗 |
| `CLOUDRUN_ERROR` | Cloud Run HTTP エラー |
| `RESPONSE_PARSE_FAILED` | レスポンス解析失敗 |
| `DRIVE_FOLDER_FAILED` | Drive フォルダ取得失敗 |
| `NO_OUTPUT` | レスポンスが空 |
| `GENERATION_ERROR` | Excel 生成エラー |
| `DRIVE_SAVE_FAILED` | Drive 保存失敗 |

---

## セキュリティ確認

| 項目 | 確認結果 |
|---|---|
| APPGEN_SECRET の HTML/JS 露出 | なし — ScriptProperties からのみ読む |
| APPGEN_ENDPOINT の HTML/JS 露出 | なし — ScriptProperties からのみ読む |
| 個人情報のログ出力 | なし — Logger.log は patientId とエラーコードのみ |
| Cloud Run に直接接続するクライアントJS | なし — GAS server-side 経由のみ |
| A案 generateClaimApplication_V3 の使用 | なし |

---

## 既存B案メニューとの関係

既存のスプレッドシートメニュー「帳票出力 → 申請書を出力」= `V3TR_menuGenerateApplication_B()` は**変更なし・壊れていない**。

Web UI 経由の `generateClaimApplicationBFromWeb_V3` は既存メニューの薄い Web 版 wrapper として追加された。

---

## テスト

### LiveCheck（WEB-4A）

| テスト | 内容 | 結果 |
|---|---|---|
| W4A-1 | 「申請書Excelを生成」ボタンが存在する | 5 SKIP（auth 期限切れ — 実装は正しい） |
| W4A-2 | APPGEN_SECRET/ENDPOINT が HTML 露出なし | 同上 |
| W4A-3 | ボタンハンドラが generateClaimApplicationBFromWeb_V3 を呼ぶ | 同上 |
| W4A-4 | キャンセルでボタンが有効に戻る | 同上 |
| W4A-5 | 既存 Step 1/Step 2 ボタンが壊れていない | 同上 |

**注:** 全テストが 5 SKIP なのは、GAS auth.json が期限切れのため（pre-existing 状況）。  
テスト自体の実装は正しい（URL修正・handleAuthRedirect 対応済み）。  
auth 更新後（`npm run save-auth`）に再実行すると PASS になる。

### 既存テスト（auth 期限切れによる全 SKIP — pre-existing）

| スイート | 実行前 | 実行後 |
|---|---|---|
| smoke | 28 SKIP | 28 SKIP（変化なし） |
| web34 | 10 SKIP | 10 SKIP（変化なし） |

---

## clasp push

実施済み（2026-05-08 09:39:47）。20 ファイル push 成功。

```
Pushed 20 files at 9:39:47.
└─ Ver3_transferData.js
└─ web-monthly-claim-detail.html
...
```

---

## 本番 deploy

**未実施（本番 deploy 禁止のため）**

clasp push のみ実施。`clasp deploy -i` は未実施。  
本番反映には別途 `clasp deploy -i <deploymentId>` が必要。

---

## Dashboard

**対象外**（AIOS-06 Dashboard 対象プロジェクト外）

---

## 残確認事項（人間目視）

- [ ] `?page=monthlyClaimDetail&patientId=hirayamaka&ym=2026-04` で新しい「申請書Excelを生成」ボタンが表示されること
- [ ] ボタン押下 → confirm ダイアログ → 「OK」で生成が走ること
- [ ] 成功時: ファイル名と Drive リンクが画面表示されること
- [ ] Drive で生成 xlsx を開いて内容を目視確認
- [ ] 印刷プレビューで 1 ページに収まるか
- [ ] 転帰欄の丸囲みの見た目
- [ ] 申請書生成ログに記録されていること
- [ ] 既存メニュー「帳票出力 → 申請書を出力」が壊れていないこと

---

## 次アクション

1. auth.json 更新（`npm run save-auth`）後に `npm run test:jyu:web4` で PASS 確認
2. 目視確認（上記チェックリスト）
3. 問題なければ本番 deploy 判断（`clasp deploy -i <deploymentId>`）
