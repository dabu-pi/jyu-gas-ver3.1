# claim-excel — 申請書 Excel 生成ツール

療養費支給申請書の Excel 帳票を生成するローカル実行ツール。

---

## 必要なもの

- Python 3.8+
- openpyxl: `pip install openpyxl`
- このディレクトリにある `application_template.xlsx`
- GAS から出力した NDJSON ファイル

---

## NDJSON の取得

Web UI で取得:

1. `?page=monthlyClaimDetail&patientId=<ID>&ym=YYYY-MM` を開く
2. 「Step 2: 申請書 NDJSON を Drive に出力」ボタンをクリック
3. 表示された Drive URL から JSON ファイルをダウンロード
4. `tools/claim-excel/` に配置する

---

## Excel 生成

```bash
cd gas-projects/jyu-gas-ver3.1/tools/claim-excel
python write_application.py <ndjson_file.json>
```

出力先: 同ディレクトリに `申請書_<ID>_<YYYY-MM>_<timestamp>.xlsx`

---

## 確認内容

生成された Excel を開いて以下を確認:

- 負傷名欄: (1)(2) に空行なく詰まっている
- 施術日カレンダー（行32）: 来院日に ○ が入っている
- 転帰: 該当箇所に丸囲みが入っている
- 金額欄: 正しい金額が入っている
- 印刷プレビュー: 1ページに収まる

---

## PDF 化

Excel を開いて「ファイル → 印刷 → PDFとして保存」または印刷で PDF 化する。
