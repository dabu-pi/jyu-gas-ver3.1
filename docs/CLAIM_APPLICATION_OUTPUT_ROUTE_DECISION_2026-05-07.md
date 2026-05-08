# 申請書出力ルート確定記録

作成日: 2026-05-07  
ステータス: **確定（本日をもって正ルート採用）**

---

## 決定事項

### 正ルート: B案 Cloud Run Excel 出力

```
スプレッドシートメニュー:
  帳票出力 → 申請書を出力

コールバック関数:
  V3TR_menuGenerateApplication_B()

ファイル:
  Ver3_transferData.js

フロー:
  GAS → NDJSON生成 → Layer2安全フィルタ → Preflight検証
  → Cloud Run /generate（X-Secret-Key認証）
  → xlsx base64 → Drive月別フォルダ保存
  → _申請書生成ログ シート記録
```

**採用理由:**
- 2026-04-06 に正式帳票で動作確認済み
- Drive 保存・申請書生成ログ記録が組み込まれている
- 2026-05-07 に hirayamaka/2026-04 で実行・人間目視確認済み

---

### 停止: Sheets直PDF (A案)

```
関数: generateClaimApplication_V3()
```

**停止理由:**
- PDF が 2 ページになる問題
- 負傷名欄の空行・残存問題
- 施術日カレンダー○印が不完全
- 転帰の丸囲みが不完全
- GAS / Sheets 直 PDF では帳票表現が不安定

**今後: 本番化しない。コードは残すが正ルートとして進めない。**

---

### 補助扱い: NDJSON + Python ローカル実行

```
exportClaimNdjson_V3
→ tools/claim-excel/write_application.py
→ application_template.xlsx
```

**位置づけ:**
- B案 Cloud Run と同一ロジック（write_application.py は完全一致）
- Cloud Run なし環境での開発・ローカル検証に使用可能
- 正ルートではない

---

## B案 実行確認結果（2026-05-07）

| 項目 | 結果 |
|---|---|
| Cloud Run URL | `https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app` |
| /health | `{"status":"ok"}` ✅ |
| / | 404 Not Found（想定内: ルートは /health と /generate のみ） |
| revision | 00026-wv2（2026-04-20 デプロイ） |
| 実行対象 | hirayamaka / 2026-04 |
| 生成ファイル | 申請書_hirayamaka_2026-04.xlsx（36,694 bytes） |
| Drive URL | https://docs.google.com/spreadsheets/d/1HHv-KH6bmfA-vEZErZ94uZxLutsxGi-C/edit |

**人間目視確認結果:**

| 確認項目 | 結果 |
|---|---|
| 負傷名（1）頸部 捻挫 | ✅ |
| 負傷名（2）腰部 捻挫 | ✅ 連続・残存なし |
| 施術日 ①⑲（4/1・4/19） | ✅ |
| 合計 4,363円 / 一部負担金 1,310円 / 請求金額 3,053円 | ✅ |
| 申請書生成ログ記録 | ✅ |

**残人間確認（印刷・目視）:**
- [ ] 印刷プレビューで 1 ページに収まるか
- [ ] 転帰欄の丸囲みの見た目
- [ ] 罫線・文字位置の最終確認

---

## GAS スクリプトプロパティ設定（B案実行に必要）

| プロパティ名 | 値 | 備考 |
|---|---|---|
| `APPGEN_ENDPOINT` | `https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app` | 末尾スラッシュなし・`/generate` 含めない |
| `APPGEN_SECRET` | Secret Manager `JREC_APPGEN_SECRET_KEY` の値 | Web/HTML/JS に絶対に出さない |

---

## Web UI からB案を呼ぶ将来方針

**現状:** Web UI から B案を呼ぶ導線は未実装（スプレッドシートメニューからのみ実行可能）  
**将来:** 以下の構成で安全に実装可能。本日は実装しない。

```
Web monthlyClaimDetail
→ google.script.run
→ GAS wrapper（V3TR_generateApplicationBCore_ の薄い呼び出し口）
→ 既存 B案ロジック（変更なし）
→ Cloud Run /generate
→ Drive 保存・ログ記録
→ Web へ Drive URL を返す
```

**制約（次回実装時に必ず守ること）:**
- Web から直接 Cloud Run を叩かない（APPGEN_SECRET 漏洩防止）
- APPGEN_SECRET を HTML/JS に出さない
- 既存 B案ロジックを壊さない（wrapper の追加のみ）

---

## 本日（2026-05-07）の作業終了状態

| 項目 | 状態 |
|---|---|
| 申請書正ルート採用 | ✅ B案 Cloud Run Excel 確定 |
| Sheets直PDF停止 | ✅ 停止・本番化しない |
| NDJSON+Python 位置づけ | ✅ 補助扱いとして記録 |
| Web からB案呼び出し | ⏸ 次回以降の実装候補として記録のみ |
| 本番 deploy | ⏸ 未実施（B案は Cloud Run のため GAS deploy 対象外） |
| 新規実装 | ⏸ 本日は記録のみで終了 |
