# Excel → PDF ルート検討記録

作成日: 2026-05-07  
背景: 申請書 PDF 帳票の品質問題（腰部捻挫重複残存）の調査

---

## 調査結果

### A案: Google Sheets テンプレート → PDF export（現行）

現在の方式。`V3TR_writeToApplication_` が「新 様式第5号」テンプレートシートに転記し、UrlFetchApp で PDF をエクスポートする。

**問題（2026-05-07 発見）:**
1. 書き込み前に旧セルをクリアしていなかった → 腰部捻挫が (3) に残存
2. `fitw=true` のみで `fith=true` がなかった → 2ページ出力

**修正状況:**
- injury row 空行詰め: ✅ 修正済み（`validInj` フィルタ）
- 旧セル clear 漏れ: ✅ 修正済み（書き込み前に全 5 行 clearContent）
- 2ページ問題: ✅ 修正済み（`fith=true` / `pagenumbers=false`）

**採用判断: 継続（修正後）**

---

### B案: Cloud Run → Excel → PDF（旧ロジック）

`V3TR_generateApplicationBCore_` が外部 Cloud Run エンドポイントに NDJSON を POST し、返ってくる base64 xlsx を Drive に保存する方式。

**現状:**
- `APPGEN_ENDPOINT` スクリプトプロパティが必要
- `APPGEN_SECRET` スクリプトプロパティが必要
- **現環境では未設定 → 利用不可**

**採用判断: 見送り（外部依存）**

---

## 決定: A案継続（修正後）

B案の Cloud Run は環境設定が必要で即時利用不可。  
A案の根本的な問題（clear 漏れ・PDF 1ページ化）は修正コードで対応完了。

次回以降 B案（Cloud Run）を導入する場合は:
- `APPGEN_ENDPOINT` を GAS スクリプトプロパティに設定
- `APPGEN_SECRET` を GAS スクリプトプロパティに設定
- `V3TR_menuGenerateApplication_B` が呼べる環境を構築

---

## 修正履歴

| 日付 | 修正内容 | ファイル |
|---|---|---|
| 2026-05-07 | injury row 空行詰め（validInj filter） | Ver3_transferData.js |
| 2026-05-07 | 書き込み前 clearContent（全5行） | Ver3_transferData.js |
| 2026-05-07 | PDF export URL に fith=true 追加 | Ver3_core.js |
