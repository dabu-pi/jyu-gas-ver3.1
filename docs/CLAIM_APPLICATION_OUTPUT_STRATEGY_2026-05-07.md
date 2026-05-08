# 申請書出力方式 設計方針

作成日: 2026-05-07  
背景: Sheets直PDF方式の帳票品質問題を調査し、旧方式と比較した結果の記録

---

## 旧方式（Python ローカル実行）の確認結果

### 旧フロー

```
GAS メニュー「一括JSON出力（月指定）」
  → Drive に NDJSON 保存（V3TR_menuBatchExportJson / V3TR_exportTransferJson_）
  → ユーザーがローカルにダウンロード
  → python write_application.py --batch <ndjson_file>
  → application_template.xlsx に openpyxl で転記
  → Excel ファイル生成（印刷設定・○印・転帰処理含む）
  → ユーザーが PDF 化（Excel の「PDFとして保存」等）
```

### 旧方式の関連ファイル

| ファイル | 内容 |
|---|---|
| `write_application.py` | Python 転記スクリプト（openpyxl 使用）|
| `application_template.xlsx` | Excel 申請書テンプレート（新 様式第5号）|
| `test_d6.xlsx` | テスト用 Excel ファイル（58KB）|

**所在:** `workspace-export/gas-projects/jyu-gas-ver3.1/`

### 旧方式が対応していた処理

| 処理 | Python 版 | GAS Sheets直PDF版 |
|---|---|---|
| 負傷名欄の空行詰め | ✅ `build_injury_rows` | ✅ 修正済み（validInj filter + clearContent）|
| 施術日カレンダー ○ 印（行32） | ✅ `put_calendar_circles` | ❌ **未実装** |
| 転帰の楕円丸囲み | ✅ ellipse 図形描画 | ❌ **未実装（GAS では図形操作困難）** |
| 保険種別の ○ 付け | ✅ `SELECTION_SPLIT_MAP` | △ 文字置換方式で対応 |
| 印刷範囲・1ページ設定 | ✅ Excel テンプレート設定 | △ URL パラメータで対応（不安定）|
| 実日数・開始終了日 | ✅ | ✅ |
| 金額欄（初検料・後療料等） | ✅ | ✅ |

---

## 経路比較

### A案: GAS Sheets テンプレート → PDF export（現行試験中）

**メリット:**
- Web UI から1クリックで完結できる（将来的に）
- Drive 保存が自動

**問題点:**
- 施術日カレンダーの ○ 印が未実装
- 転帰の楕円丸囲みが未実装（GAS では Slides API 等が必要）
- PDF 1ページ化の安定性に懸念
- テンプレートセルのクリア処理が複雑

**現状: 試験的（帳票完全対応前）**

---

### B案: NDJSON 出力 + Python スクリプト（推奨）

**メリット:**
- 既存の完成済みロジックを再利用
- 施術日カレンダー ○・転帰・印刷設定がすでに実装されている
- 帳票品質が保証されている

**手順:**
1. Web UI の「申請書 NDJSON を Drive に出力」ボタンをクリック
2. Drive から NDJSON をダウンロード
3. ローカルで `python write_application.py --batch <ファイル名>` を実行
4. 生成された Excel ファイルを PDF 化（印刷 → PDF 保存）

**制約:**
- Python 環境が必要（`pip install openpyxl`）
- `application_template.xlsx` がローカルに必要
- ネットワーク上で完全自動化するには Cloud Run（B案 API）が必要

---

### C案: Cloud Run API（B案 API）

`APPGEN_ENDPOINT` と `APPGEN_SECRET` を設定すれば、GAS から Cloud Run 経由で自動 xlsx 生成が可能。

**現状: APPGEN_ENDPOINT 未設定のため利用不可**

---

## 採用方針

| フェーズ | 採用方式 | 理由 |
|---|---|---|
| 現在（2026-05）| **B案（NDJSON + Python）** | 帳票品質保証・既存完成ロジック再利用 |
| 将来 | A案の完成（施術日○・転帰を追加実装） または C案（Cloud Run 設定） | Web UI 完結のため |

---

## GAS 実装状況（2026-05-07）

| 機能 | 状況 |
|---|---|
| `exportClaimNdjson_V3(pid, ym)` | ✅ 実装完了 |
| `generateClaimApplication_V3(pid, ym)` | △ 試験的（施術日○・転帰未対応） |
| web-monthly-claim-detail.html | ✅ NDJSON 出力ボタン + PDF 試験ボタンを併設 |

---

## 本番 deploy 方針

```
DEPLOY: 未実施
理由:
  - 施術日カレンダー ○・転帰処理がA案に未実装
  - B案（NDJSON + Python）での帳票確認を優先する
  - A案が帳票として完成した後に deploy を判断する
```
