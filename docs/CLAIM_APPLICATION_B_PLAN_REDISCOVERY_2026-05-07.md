# 申請書生成B案 再確認記録

作成日: 2026-05-07  
ステータス: **B案採用確定 — スプレッドシートメニューから実行・人間目視確認済み**

---

## ユーザー指摘（2026-05-07）

> 完成度が一番高かった申請書出力は、スプレッドシートのメニューから実行していた「申請書生成B案」です。
> これは Excel で出力されていたもので、これまでの確認では最も成功していた申請書生成ルートです。

この指摘を受けて、B案の実体・現在の状態・実行可能性を調査した。

---

## 調査結果

### B案メニュー定義

| 項目 | 内容 |
|---|---|
| メニュー構成 | 柔整ツール > 帳票出力 > **申請書を出力** |
| コールバック関数 | `V3TR_menuGenerateApplication_B()` |
| ソースファイル | `Ver3_transferData.js` (行 2509〜2872) |
| ダイアログタイトル | 申請書生成（B案） |

### B案フロー

```
1. スプレッドシートメニュー「帳票出力」→「申請書を出力」
2. ダイアログ表示（患者ID省略可・対象月入力）
3. V3TR_runGenerateApplicationDialog(pid, ym)
4. V3TR_generateApplicationBCore_(patientIds, ym)
   ├── NDJSON生成（V3TR_buildTransferDataForMonth_ + V3TR_exportTransferJson_）
   ├── Layer 2 安全フィルタ（claimPay=0 患者を除外）
   ├── Preflight validation（必須キー・対象月一致・金額整合）
   ├── Cloud Run POST /generate（X-Secret-Key認証）
   ├── レスポンス base64 → Drive 月別フォルダに xlsx 保存
   └── _申請書生成ログ シートへ記録
```

### Cloud Run サービス

| 項目 | 値 |
|---|---|
| URL | `https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app` |
| GCPプロジェクト | `hirayama-jrec-appgen` |
| リージョン | `asia-northeast1`（東京） |
| 現行リビジョン | `00026-wv2` |
| 最終デプロイ | 2026-04-20 |
| /health 確認 | `{"status":"ok"}` ✅（2026-05-07 確認） |

### B案実行結果（2026-05-07 実行）

対象: hirayamaka / 2026-04

| 項目 | 結果 |
|---|---|
| status | ok |
| 生成ファイル | `申請書_hirayamaka_2026-04.xlsx` |
| ファイルサイズ | 36,694 bytes |
| E26 負傷名 | `（1）頸部 捻挫` ✅ |
| カレンダー○ | Pillow 画像埋込（目視確認が必要） |
| 金額 | ¥3,053（目視確認が必要） |

### スプレッドシートメニューからの実行・人間目視確認（2026-05-07）

GAS スクリプトプロパティに `APPGEN_ENDPOINT` / `APPGEN_SECRET` が設定されており、
スプレッドシートメニュー「帳票出力 → 申請書を出力」からB案を実行した。

**生成ファイル Drive URL:**  
https://docs.google.com/spreadsheets/d/1HHv-KH6bmfA-vEZErZ94uZxLutsxGi-C/edit

**目視確認結果:**

| 確認項目 | 結果 |
|---|---|
| 負傷名（1）頸部 捻挫 | ✅ 正しく入力されている |
| 負傷名（2）腰部 捻挫 | ✅ 正しく連続している |
| （3）以降に不要な残存 | ✅ なし |
| 施術日 ①⑲（4/1・4/19） | ✅ 入力されている |
| 合計 4,363円 | ✅ 確認済み |
| 一部負担金 1,310円 | ✅ 確認済み |
| 請求金額 3,053円 | ✅ 確認済み |
| 申請書生成ログ記録 | ✅ 本日分が記録されている |

**残人間確認（目視・印刷）:**
- [ ] Excel 印刷プレビューで 1 ページに収まるか
- [ ] 転帰欄の丸囲みの見た目
- [ ] 罫線・文字位置の最終確認

---

## B案と NDJSON+Python ローカル実行の関係

| 方式 | 用途 | 位置づけ |
|---|---|---|
| B案（Cloud Run経由） | 正式帳票出力（Drive保存・ログ記録付き） | ✅ **正ルート** |
| ローカル実行（tools/claim-excel/） | 開発・デバッグ・Cloud Runなし環境 | 補助ルート |

**同一ロジック**: B案のCloud Runが内部で `write_application.py` を実行している。
`tools/claim-excel/write_application.py` と `workspace-export/gas-projects/jyu-gas-ver3.1/write_application.py` は完全一致（1719行）。

---

## B案実行に必要なGAS設定

GAS スクリプトプロパティに以下を設定すること:

| プロパティ名 | 値 |
|---|---|
| `APPGEN_ENDPOINT` | `https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app` |
| `APPGEN_SECRET` | Secret Manager `JREC_APPGEN_SECRET_KEY` の値 |

設定手順:
1. スプレッドシート > 拡張機能 > Apps Script
2. 左メニュー「プロジェクトの設定」>「スクリプトプロパティ」
3. 上記2項目を追加/更新

---

## 各ルートの位置づけ（確定）

| ルート | 位置づけ | 理由 |
|---|---|---|
| **B案（Cloud Run）** | ✅ **正ルート** | 2026-04-06 正式帳票で確認済み・Drive保存・ログ記録 |
| Sheets直PDF（A案） | ❌ 停止 | 施術日カレンダー○・転帰が未実装で帳票不完全 |
| NDJSON + Python ローカル | 補助 | Cloud Run の代替として開発時使用可能 |

---

## Cloud Run エンドポイント補足

| パス | 説明 |
|---|---|
| `/` | 404 Not Found — **想定内**。ルートは定義されていない |
| `/health` | `{"status":"ok"}` — 稼働確認用 |
| `/generate` | POST — xlsx 生成 API（GAS側で `endpoint + "/generate"` を付与） |

**APPGEN_ENDPOINT の設定値:**
```
https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app
```
末尾に `/generate` は含めない。GAS側の `UrlFetchApp.fetch(endpoint + "/generate", options)` で付与される。

---

## Web UI からB案を呼ぶ将来方針（次回以降の候補）

現時点では、Web UI からB案（Cloud Run xlsx 生成）を呼ぶ導線は未実装。
技術的には可能。以下の構成で安全に実装できる。

```
Web monthlyClaimDetail
→ google.script.run
→ GAS側 wrapper（既存B案ロジックの呼び出し口）
→ V3TR_generateApplicationBCore_()
→ Cloud Run /generate
→ xlsx 生成・Drive 保存・ログ記録
→ Web へ Drive URL を返す
```

**実装上の制約（必ず守ること）:**

| 制約 | 理由 |
|---|---|
| Web から直接 Cloud Run を叩かない | APPGEN_SECRET が Web に漏れる |
| APPGEN_SECRET を HTML/JS に出さない | 秘密情報の漏洩防止 |
| 既存 B案ロジックを壊さない | スプレッドシートメニューからの運用を維持 |
| wrapper は既存B案の薄い呼び出し口として追加 | 既存コードの大幅変更は禁止 |

**本日は実装しない。次回以降の実装候補として記録する。**

---

## 残人間確認

| 確認項目 | 内容 |
|---|---|
| GAS設定確認 | APPGEN_ENDPOINT / APPGEN_SECRET の設定状況（スクリプトエディタで確認） ✅ 設定済み（実行確認済み） |
| Excel目視確認 | 印刷プレビュー1ページ / 転帰欄の丸囲み / 罫線・文字位置 |
| Cloud Run再デプロイ判断 | workspace-export の write_application.py に変更があれば再デプロイ検討 |

---

## 関連ファイル

| ファイル | 内容 |
|---|---|
| `Ver3_transferData.js:2494` | B案メニュー関数本体 |
| `Ver3_core.js:413` | メニュー定義（帳票出力 > 申請書を出力） |
| `workspace-export/.../server.py` | Cloud Run Flaskサーバー |
| `workspace-export/.../Dockerfile` | Cloud Runイメージビルド定義 |
| `tools/claim-excel/write_application.py` | xlsx生成ロジック（Cloud Runと同一） |
| `docs/JREC-01_CloudRun_デプロイ手順.md` | Cloud Runデプロイ手順書（workspace-export内） |
