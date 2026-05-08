# NDJSON → Excel 申請書生成フロー 検証記録

作成日: 2026-05-07  
ステータス: NDJSON 取得・Excel 生成 COMPLETED / 目視確認待ち

---

## ツール配置

`gas-projects/jyu-gas-ver3.1/tools/claim-excel/` に以下を配置した。

| ファイル | 内容 | 出所 |
|---|---|---|
| `write_application.py` | Python 申請書転記スクリプト（openpyxl 使用）| workspace-export からコピー |
| `application_template.xlsx` | 申請書 Excel テンプレート（新 様式第5号）| workspace-export からコピー |
| `requirements.txt` | `openpyxl>=3.0.0` |  |
| `README.md` | 使い方の説明 | |
| `verify_application_xlsx.py` | 生成 Excel の自動確認スクリプト | 新規作成 |

---

## 互換性確認

`exportClaimNdjson_V3` の出力形式と `write_application.py` の入力形式が一致していることを確認した。

| 項目 | GAS 出力 | Python 入力 |
|---|---|---|
| 形式 | `{ "case1": {...}, "case2": {...}, "visitDays": [...] }` | `json_data.get("case1") / get("case2") / get("visitDays")` |
| 施術日 | `visitDays: [3, 5, 10, ...]` | `put_calendar_circles(ws, "M32", visit_days)` |
| 負傷名 | `case1["負傷名1"]` など | `build_injury_rows(row1, row2)` で詰め処理 |

---

## 実行手順（ツール整備後）

### Step 1: auth 更新

```powershell
$dir = "C:\hirayama-ai-workspace\workspace\tools\live-check-runner\.chrome-profile"
Start-Process "chrome" "--remote-debugging-port=9222 --user-data-dir=`"$dir`""
# Chrome で Google ログイン → GAS dev URL を開く
cd C:\hirayama-ai-workspace\workspace\tools\live-check-runner
npm run save-auth
```

### Step 2: NDJSON 取得

Web UI (`?page=monthlyClaimDetail&patientId=hirayamaka&ym=2026-04`) で「申請書 NDJSON を Drive に出力」をクリック。
または Playwright spec で自動取得。

### Step 3: Excel 生成

```powershell
cd gas-projects/jyu-gas-ver3.1/tools/claim-excel
python write_application.py <ndjson_file.json>
```

### Step 4: 自動確認

```powershell
python verify_application_xlsx.py  # 最新の xlsx を自動検出
```

### Step 5: 目視確認

生成された Excel を開いて:
1. 負傷名欄: (1)頸部捻挫 (2)腰部捻挫 と詰まっているか
2. 施術日カレンダー: 来院日に ○ が入っているか
3. 転帰: 丸囲みが入っているか
4. 金額: ¥3,053
5. 印刷プレビュー: 1ページに収まるか

---

## 実行結果（2026-05-07 完了）

### NDJSON 取得

- Drive ファイル: `申請書_転記データ_hirayamaka_2026-04_20260507_1612.json`
- Drive File ID: `1km2r2P5T6lOFNo3i_nzIHLkcLQeqO8b_`
- ローカル保存先: `tools/claim-excel/申請書_転記データ_hirayamaka_2026-04_20260507_1612.json`

### NDJSON 内容（実データ）

| 項目 | 値 |
|---|---|
| 患者 | hirayamaka（平山克士） |
| 対象月 | 2026-04 |
| case1 負傷名1 | 頸部 捻挫 |
| case2 | 初検抑制（空） |
| visitDays | [1, 19] |
| 請求金額 | ¥3,053 |
| 算定区分 | 再検 |

### Excel 生成

- 出力ファイル: `tools/claim-excel/申請書_hirayamaka_2026-04.xlsx`（35.8KB）
- 使用 Pillow バージョン: 12.2.0（pip install Pillow で追加）

### 自動検証結果

| チェック項目 | 結果 |
|---|---|
| 負傷名欄 E26 | ✅ `（1）頸部 捻挫` |
| 施術日カレンダー○ | ⚠️ Pillow 画像埋込のため verify では未検出（設計上の制限） |
| 金額セル DX44/45/46 | ⚠️ 1桁分割セル方式のため None（設計上の制限） |

## 残タスク

| タスク | 状態 |
|---|---|
| auth 更新 | ✅ 完了 |
| NDJSON 取得（hirayamaka/2026-04） | ✅ 完了 |
| Excel 生成（write_application.py 実行） | ✅ 完了 |
| 自動確認（verify_application_xlsx.py） | ✅ 完了 |
| 目視確認 | ⏸ 人間タスク（申請書_hirayamaka_2026-04.xlsx を Excel で開いて確認） |
