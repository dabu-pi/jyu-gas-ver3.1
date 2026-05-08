# WEB-2 LiveCheck 記録

実施日: 2026-05-06  
実施者: Claude Code  
対象: Phase WEB-2 実装内容（来院記録登録 MVP）  
clasp push: 完了（15ファイル）

---

## Playwright LiveCheck 実行結果（2026-05-06）

**スペックファイル:** `tools/live-check-runner/projects/jyu-gas-ver31/web2.spec.ts`  
**実行コマンド:** `npm run test:jyu:web2`

| 結果 | 件数 |
|---|---|
| PASS | 0 |
| FAIL | 0 |
| SKIP | 16（全件） |

**SKIP 理由:** Google Account Chooser に遷移（WEB-1 と同じ auth 問題）

### 最終結果（testData.patientId 設定・全テスト実行）★

| 結果 | 件数 |
|---|---|
| **PASS** | **16（全件）** |
| FAIL | **0** |
| SKIP | **0** |

**npm run test:jyu 全体: 42 passed / 0 failed / 0 skipped**

**WEB-2 PASS 確認項目（web2.spec.ts）:**

| テスト | chromium | 確認内容 |
|---|---|---|
| W2-1a: page=visitNew 到達 (HTTP<400) | ✅ | ページ到達 |
| W2-1b: ページタイトルに「来院記録」 | ✅ | doGet page=visitNew 正常 |
| W2-2a: #visitDate 存在 | ✅ | 来院日入力欄 |
| W2-2b: #accountingType 存在 | ✅ | 会計区分セレクト |
| W2-2c: #kubun 存在 | ✅ | 区分セレクト |
| W2-2d: #bodyPart 存在 | ✅ | 部位入力欄 |
| W2-2e: #disease 存在 | ✅ | 傷病名入力欄 |
| W2-2f: #injuryDate 存在 | ✅ | 受傷日入力欄 |
| W2-2g: .btn-save 存在 | ✅ | 「来院を登録する」ボタン |
| W2-3a: patientId なし → #patient-chip「未指定」 | ✅ | patientId 引き継ぎ |
| W2-3b: patientId あり → chip 表示 | ✅ | patientId が chip に表示される |
| W2-4a: #inheritBtn 存在 | ✅ | 前回引き継ぎボタン |
| W2-4b: 前回引き継ぎ応答 | ✅ | getPrevVisitData_V3 正常応答 |
| W2-5: 必須未入力 → モーダル開かない | ✅ | バリデーション動作確認 |
| W2-6/7: 必須入力 → モーダル → キャンセル | ✅ | 確認モーダル開閉確認 |
| W2-8: コンソールエラーなし | ✅ | 重大エラーなし確認 |

**W2-4b 実行ログ（前回来院データあり確認）:**
```
[W2-4b] dialog: 前回来院（2026-04-19）のデータを引き継ぎました。内容を確認・修正してください。
```
→ `getPrevVisitData_V3` が来院ケースシートから前回データを正常取得。

**W2-5 実行ログ（バリデーション確認）:**
```
[W2-5] validation alert: "患者IDが未指定です。"
```
→ 必須項目未入力時に `alert()` が発火し、モーダルが開かないことを確認。

---

## 確認方式

コードレビュー LiveCheck。実機ブラウザ確認は現場スマホで実施すること。

---

## 1. 実装ファイル一覧

| 変更種別 | ファイル | 内容 |
|---|---|---|
| 追記 | `Ver3_core.js` | `getPrevVisitData_V3`, `findLatestCaseKeyForPatient_`, `saveVisitFromWeb_V3` を末尾追加 |
| 追記 | `Ver3_core.js` | `doGet` に `page=visitNew` ハンドラ追加 |
| 新規 | `web-visit-new.html` | 来院記録登録フォーム（WEB-2 新規） |
| 変更 | `web-patient-detail.html` | 「来院記録を追加」ボタン追加 |

---

## 2. Ver3_core.js 追加関数

### getPrevVisitData_V3 (5831行〜)

| 確認項目 | 結果 |
|---|---|
| 来院ケースシートから最新 treatDate の行を取得 | ✅ treatDate 降順ソート |
| 複数ケース（caseNo=1/2）対応 | ✅ latestRows で同日複数行を取得 |
| 返却データ（kubun/bodyPart/disease/injuryDate/cold/warm/elec） | ✅ 正常 |
| 個人情報ログ禁止 | ✅ patientId + 件数のみ |
| エラーハンドリング | ✅ `{ ok: false, error: ... }` |

### findLatestCaseKeyForPatient_ (5924行〜)

| 確認項目 | 結果 |
|---|---|
| 患者ID + caseNo で絞り込み | ✅ 正常 |
| treatDate 降順で最新 caseKey を返す | ✅ 正常 |
| caseKey 列が存在しない場合の null 返却 | ✅ 正常 |

### saveVisitFromWeb_V3 (5976行〜)

| 確認項目 | 結果 |
|---|---|
| patientId バリデーション | ✅ 空欄で MISSING_REQUIRED |
| visitDate 形式バリデーション（YYYY-MM-DD） | ✅ 正規表現チェック |
| cases 空配列バリデーション | ✅ MISSING_REQUIRED |
| 患者マスタ存在確認 | ✅ 正常 |
| buildVisitKey_ 再利用 | ✅ 既存関数を呼ぶ |
| 二重登録チェック（来院ヘッダ） | ✅ DUPLICATE_VISIT |
| 初検 → episodeStartDate = visitDate | ✅ 正常 |
| 再検/後療 → findLatestCaseKeyForPatient_ から解析 | ✅ 正常 |
| 来院ケース行を appendRow | ✅ caseSh.appendRow(rowArr) |
| 来院ヘッダを appendHeaderRow_V3_ で保存 | ✅ 正常 |
| visitTotal = 0 / needCheck = true | ✅ 金額未算定フラグ |
| needCheckReason = "Web UI 登録（金額未算定）" | ✅ 正常 |
| ログ方針（個人情報なし） | ✅ patientId/visitKey/result のみ |
| 返却: `{ok:true, visitKey, patientId, message}` | ✅ 正常 |
| エラー: `{ok:false, reasonCode, message}` | ✅ 正常 |

### doGet `page=visitNew` (5452行〜)

| 確認項目 | 結果 |
|---|---|
| `page=visitNew` → `web-visit-new.html` | ✅ 正常 |
| `patientId` パラメータをテンプレートに注入 | ✅ `tmplVisit.patientId = vPid` |
| `appBaseUrl` 注入 | ✅ 正常 |

---

## 3. web-visit-new.html（コードレビュー）

| 確認項目 | 結果 |
|---|---|
| 戻りリンク: patientId あり → `?page=detail&patientId=xxx` | ✅ 正常 |
| 戻りリンク: patientId なし → `?page=search` | ✅ 正常 |
| 来院日: 今日の日付を初期値にセット | ✅ 正常 |
| 「前回引き継ぎ」→ `getPrevVisitData_V3` 呼び出し | ✅ 正常 |
| フォームバリデーション（保存前）: 患者ID/来院日/区分/部位/傷病 必須 | ✅ 正常 |
| 確認モーダルで入力内容を表示 | ✅ 正常 |
| `saveVisitFromWeb_V3(payload)` 呼び出し | ✅ 正常 |
| 保存成功 → visitKey 表示 + 患者詳細へのリンク | ✅ 正常 |
| 保存失敗 → エラーメッセージ表示 | ✅ 正常 |
| XSS 対策（`esc()` 関数） | ✅ 正常 |
| `window.location` 使用禁止 | ✅ 全て `APP_BASE_URL` ベース |
| `google.script.run` のローディング表示 | ✅ `loadingOverlay` で保存中を表示 |

---

## 4. web-patient-detail.html 変更（コードレビュー）

| 確認項目 | 結果 |
|---|---|
| 「来院記録を追加 →」ボタン追加 | ✅ `btn-visit-new` |
| ボタンURL: `?page=visitNew&patientId=xxx` | ✅ 正常 |
| 「自費明細入力 →」ボタン（既存）の動作維持 | ✅ `btn-selfpay` href は変更なし |
| 両ボタンは `action-row` でグループ化 | ✅ `display:flex` で横並び |

---

## 5. 既存稼働導線への影響

| 既存機能 | 影響 |
|---|---|
| `page=search` (patientSearch.html) | なし（変更なし） |
| `page=selfpay` (selfPayWeb.html) | なし（変更なし） |
| `page=home` (web-home.html) | なし（変更なし） |
| `saveVisit_V3()` スプレッドシートUI保存 | なし（既存関数は変更なし） |
| `getPatientDetail_V3()` | なし（変更なし） |

---

## 6. WEB-2 スコープ外（後フェーズ）

| 機能 | 理由 | フェーズ |
|---|---|---|
| 保険金額自動計算 | calcVisitAmounts_V3_ のシート依存が深い | WEB-2.5 |
| ケース2入力 | MVP に不要 | WEB-2.5 |
| 施術終了日・転帰入力 | MVP に不要 | WEB-2.5 |
| keikaNow / shoken 入力 | セル結合の複雑さ | WEB-2.5 |
| 登録後 needCheck 解消（金額算定） | スプレッドシートで実施 | WEB-2.5 |

---

## 7. 要実機確認（現場スマホ）

1. **`?page=visitNew&patientId=実在患者ID`** → フォームが表示されるか
2. **「前回引き継ぎ」ボタン** → 前回データが入力欄にセットされるか
3. **フォーム入力 → 確認モーダル** → 入力内容が正しく表示されるか
4. **「登録する」ボタン** → 来院ケース / 来院ヘッダに行が追加されるか
5. **来院ヘッダの 要確認 = TRUE / 来院合計 = 0** であることを確認
6. **重複登録（同日同患者）** → DUPLICATE_VISIT エラーが表示されるか
7. **患者詳細ページの「来院記録を追加」ボタン** → visitNew に遷移するか

---

## 8. 総合判定

| 機能 | 判定 |
|---|---|
| `getPrevVisitData_V3` | ✅ PASS（コードレビュー） |
| `saveVisitFromWeb_V3` | ✅ PASS（コードレビュー） |
| `page=visitNew` ルーティング | ✅ PASS（コードレビュー） |
| `web-visit-new.html` | ✅ PASS（コードレビュー） |
| `web-patient-detail.html` 変更 | ✅ PASS（コードレビュー） |
| 既存導線への影響 | ✅ なし |
| clasp push | ✅ 完了（15ファイル） |
