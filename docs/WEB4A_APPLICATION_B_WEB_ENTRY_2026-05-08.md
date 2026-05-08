# WEB-4A: Web UI B案申請書生成入口 実装記録

作成日: 2026-05-08  
ステータス: **実機確認完了 / clasp push 済 / commit & push 済 / 本番 deploy 未実施**

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

### LiveCheck（WEB-4A）— 2026-05-08 auth 更新後実行

| テスト | 内容 | 結果 |
|---|---|---|
| W4A-1 | 「申請書Excelを生成」ボタンが存在する | ✅ PASS |
| W4A-2 | APPGEN_SECRET/ENDPOINT が HTML 露出なし | ✅ PASS |
| W4A-3 | ボタンハンドラが generateClaimApplicationBFromWeb_V3 を呼ぶ | ✅ PASS（テスト修正: 否定チェック削除） |
| W4A-4 | キャンセルでボタンが有効に戻る | ✅ PASS |
| W4A-5 | 既存 Step 1/Step 2 ボタンが壊れていない | ✅ PASS |

**W4A-3 修正メモ:** `not.toContain("generateClaimApplication_V3(PATIENT_ID")` が Step 4(PDF) の `doPdfGenerate` 内の合法的な呼び出しを誤検知した。正の確認（`doAppgenB` / `generateClaimApplicationBFromWeb_V3` 含有）のみに変更。

### 実生成確認（2026-05-08 10:29:43）

| 項目 | 結果 |
|---|---|
| 対象 | hirayamaka / 2026-04 |
| 結果 | ✅ 成功 |
| FileName | `申請書_hirayamaka_2026-04_102943.xlsx` |
| Drive保存 | ✅（申請書生成ログに記録済み） |
| confirm ダイアログ | 正常表示・OK押下で生成開始 |
| 成功メッセージ | ✅ 表示（Drive リンク付き） |

### 回帰テスト（auth 更新後）

| スイート | 結果 |
|---|---|
| smoke | **28 PASS** |
| web3 | **8 PASS** |
| web34 | **9 PASS / 1 SKIP**（W3.4-10 cleanup dry-run — 設計通り） |
| web4 (WEB-4A) | **5 PASS** |

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

- [x] `?page=monthlyClaimDetail&patientId=hirayamaka&ym=2026-04` でボタンが表示されること ✅ LiveCheck W4A-1 PASS
- [x] ボタン押下 → confirm ダイアログ → 「OK」で生成が走ること ✅ 実生成確認済み
- [x] 成功時: ファイル名と Drive リンクが画面表示されること ✅ 実生成確認済み
- [x] 申請書生成ログに記録されていること ✅ 実生成確認済み
- [ ] Drive で生成 xlsx `申請書_hirayamaka_2026-04_102943.xlsx` を開いて内容を目視確認（**人間タスク**）
- [ ] 印刷プレビューで 1 ページに収まるか（**人間タスク**）
- [ ] 転帰欄の丸囲みの見た目（**人間タスク**）
- [ ] 罫線・文字位置の最終確認（**人間タスク**）
- [ ] 既存メニュー「帳票出力 → 申請書を出力」が壊れていないこと（スプレッドシート上で確認 / **人間タスク**）

---

## 次アクション

1. Drive で `申請書_hirayamaka_2026-04_102943.xlsx` を開き、目視・印刷プレビュー確認（人間タスク）
2. 確認 OK → 本番 deploy 判断（`clasp deploy -i <deploymentId>`）
3. 本番 deploy 後: `/exec?page=monthlyClaimDetail` でも同じ動作を確認
