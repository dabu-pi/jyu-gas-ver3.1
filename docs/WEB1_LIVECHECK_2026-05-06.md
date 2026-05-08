# WEB-1 LiveCheck 記録

実施日: 2026-05-06  
実施者: Claude Code  
対象: Phase WEB-1 実装内容 + 既存稼働導線

---

## Playwright LiveCheck 実行結果（2026-05-06 第1回・第2回）

**スペックファイル:** `tools/live-check-runner/projects/jyu-gas-ver31/smoke.spec.ts`  
**実行コマンド:** `npm run test:jyu:smoke`

### 第1回（Account Chooser 問題）
| 結果 | 件数 |
|---|---|
| PASS | 0 |
| FAIL | 0 |
| SKIP | 26（全件） |
SKIP 理由: auth.json に JYU-GAS セッションなし

### 第2回（auth 更新後・RTS 期限切れ判明）
| 結果 | 件数 |
|---|---|
| PASS | 0 |
| FAIL | 0 |
| SKIP | 26（全件）・JREC-SF01 も同様 |
SKIP 理由: `__Secure-1PSIDRTS` / `__Secure-3PSIDRTS` 期限切れ（2026-05-04 01:08 UTC）

### 第4回（testData.patientId 設定・S-6 修正後）★最終
| 結果 | 件数 |
|---|---|
| **PASS** | **42（全件）** |
| FAIL | **0** |
| SKIP | **0** |

**WEB-1 PASS 確認項目（smoke.spec.ts）:**

| テスト | chromium | mobile |
|---|---|---|
| S-1a: devUrl 到達 (HTTP<400) | ✅ | ✅ |
| S-1b: page=search — h1「患者検索」 | ✅ | ✅ |
| S-1c: page=search — #keyword 存在 | ✅ | ✅ |
| S-1d: page=search — #searchBtn 存在 | ✅ | ✅ |
| S-2a: page=home 到達 | ✅ | ✅ |
| S-2b: page=home — 「JREC-01」テキスト | ✅ | ✅ |
| S-2c: page=home — 「患者検索」カード | ✅ | ✅ |
| S-3: デフォルト URL は page=search | ✅ | ✅ |
| S-4a: モバイル水平スクロールなし (search) | ✅ | ✅ |
| S-4b: モバイル水平スクロールなし (home) | ✅ | ✅ |
| S-5: page=detail (patientId なし) graceful | ✅ | ✅ |
| S-6: page=detail — 患者情報表示 | ✅ | ✅ |
| S-6b: 「来院記録を追加」ボタン存在 | ✅ | ✅ |

**testData.patientId:** 検証用実在ID（find-patient-id.ts スクリプトで取得）を config.json に設定済み。

---

## auth 失敗の根本原因（2026-05-06 診断）

### 診断結果

```
実行: tsx scripts/diag-jyu-auth.ts
確認パターン: dev/exec × normal/stealth × authuser=0/1 — 全 5 パターン FAIL
全パターン: accounts.google.com/v3/signin/accountchooser にリダイレクト
影響範囲: JYU-GAS と JREC-SF01 の両方が SKIP（auth 問題はプロジェクト横断）
```

### 期限切れクッキー

| クッキー名 | domain | 期限 | 状態 |
|---|---|---|---|
| `__Secure-1PSIDRTS` | `.google.com` | 2026-05-04 01:08 UTC | **期限切れ** |
| `__Secure-3PSIDRTS` | `.google.com` | 2026-05-04 01:08 UTC | **期限切れ** |

`PSIDRTS` = Google セッションのローテーショントークン（有効期間: 約24時間）。  
このトークンが期限切れになると、セッションが再検証のために Account Chooser を表示する。

### 正しい save-auth 手順（RTS 更新を含む）

**手順の重要ポイント:**
- Chrome を起動したあと、**必ず Google のページをアクティブに開いて**  
  RTS が更新されてから save-auth を実行する
- save-auth の直前に Chrome で Google ページが開いていることを確認すること

```powershell
# === save-auth 正しい手順 ===

# 1. Chrome を remote debugging で起動
$dir = "C:\hirayama-ai-workspace\workspace\tools\live-check-runner\.chrome-profile"
Start-Process "chrome" "--remote-debugging-port=9222 --user-data-dir=`"$dir`""

# 2. Chrome で以下を順番に開く（ログイン状態の確認 + RTS 更新）
#    a) https://accounts.google.com   ← まずここでセッション更新
#    b) pinshanka24@gmail.com でログイン確認
#    c) https://script.google.com/macros/s/AKfycbzj47fbRvTlVixUrUiV_25xkevfyI_HXhFaBKYodB2B/dev
#       → Account Chooser が出たら pinshanka24@gmail.com を選択
#       → GAS ページが表示されるまで待つ（重要）

# 3. GAS ページが表示されたのを確認してから save-auth 実行
cd C:\hirayama-ai-workspace\workspace\tools\live-check-runner
npm run save-auth

# 4. テスト実行（JYU + JREC の両方を確認）
npm run test:jyu
npm run test:jrec:smoke
```

**期待結果（save-auth 完了後）:** 42 PASS / 0 FAIL / 0 SKIP（testData.patientId 設定分は SKIP）

---

## 確認方式

GAS Web App はブラウザから直接アクセスできないため、  
**コードレビュー LiveCheck** を実施した。  
実機ブラウザ確認は現場スマホで実施すること（下記「要実機確認」参照）。

---

## 1. 作業場所・ブランチ確認

| 項目 | 結果 |
|---|---|
| `git rev-parse --show-toplevel` | `C:/hirayama-ai-workspace/workspace` |
| 現在ブランチ | `feature/auto-dev-phase3-loop` |
| リモートとの差分 | up to date（pull 済み） |

---

## 2. clasp deployments 確認

```
Found 9 deployments.
@HEAD / @7 / @8 / @9（最新番号）/ @3(x2) / @4 / @5 / @6
```

- 最新バージョン: `@9`（`AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_...`）
- Web App URL: `https://script.google.com/macros/s/AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_Ilt8SS5P5zodfF2dnmKeqso8BL8hcinVEBrQ/exec`
- `appsscript.json` の `access: "MYSELF"` → スクリプトオーナーのみアクセス可

---

## 3. doGet ルーティング（Ver3_core.js:5445-5490）

| ルート | HTMLファイル | 確認結果 |
|---|---|---|
| デフォルト（page未指定 / page=search） | `patientSearch.html` | ✅ 正常（既存と変更なし） |
| `page=home` | `web-home.html` | ✅ 正常（WEB-1 追加） |
| `page=detail&patientId=xxx` | `web-patient-detail.html` | ✅ 正常（WEB-1 追加） |
| `page=selfpay` | `selfPayWeb.html` | ✅ 正常（既存と変更なし） |

- 全ページに `appBaseUrl`（ScriptApp.getService().getUrl()）を注入済み ✅
- `page=detail` は patientId 未指定時「patients ID が指定されていません」エラーを HTML 側で表示 ✅

---

## 4. 既存稼働導線 — patientSearch.html（コードレビュー）

| 確認項目 | 結果 |
|---|---|
| `searchPatients_V3(keyword)` 呼び出し | ✅ 正常 |
| `setPatientAndDate_V3(patientId)` 呼び出し | ✅ 正常 |
| `自費明細入力 →` リンク（selfpayUrl）生成 | ✅ `APP_BASE_URL + "?page=selfpay&visitKey=" + encodeURIComponent(visitKey)` |
| `患者詳細を見る →` ボタン（detailUrl）生成 | ✅ `APP_BASE_URL + "?page=detail&patientId=" + encodeURIComponent(patientId)`（WEB-1追加） |
| `← Web ホームへ` リンク | ✅ `APP_BASE_URL + "?page=home"`（WEB-1追加） |
| `window.location` 使用禁止 | ✅ 使っていない（全て APP_BASE_URL ベース） |

---

## 5. WEB-1 新規ページ — web-home.html（コードレビュー）

| 確認項目 | 結果 |
|---|---|
| `患者検索` カード → `?page=search` リンク | ✅ 正常 |
| 未実装機能のカード | ✅ `disabled`（来院記録: Phase WEB-2、施術録/申請書: Phase WEB-3） |
| `appBaseUrl` テンプレート注入 | ✅ `<?= appBaseUrl ?>?page=search` 形式 |

---

## 6. WEB-1 新規ページ — web-patient-detail.html（コードレビュー）

| 確認項目 | 結果 |
|---|---|
| `getPatientDetail_V3(PATIENT_ID)` 呼び出し | ✅ window.onload で呼び出し |
| patientId 未指定時のエラーハンドリング | ✅ `showError("患者IDが指定されていません。")` |
| 患者基本情報表示（name/furi/birthday） | ✅ 正常 |
| 来院履歴テーブル（最大10件） | ✅ treatDate 降順、全列表示 |
| 「自費明細入力 →」ボタン（今日の visitKey） | ✅ `APP_BASE_URL + "?page=selfpay&visitKey=" + todayVK` |
| `← 患者検索に戻る` リンク | ✅ `?page=search` |
| XSS 対策（`esc()` 関数） | ✅ &/</>/& を全エスケープ |

---

## 7. getPatientDetail_V3（Ver3_core.js:5808-5880）

| 確認項目 | 結果 |
|---|---|
| `SHEETS.master`（"患者マスタ"）検索 | ✅ 正常 |
| `SHEETS.header`（"来院ヘッダ"）検索 | ✅ 正常 |
| HEADER_COLS との列名一致 | ✅ 全列名一致（visitKey/施術日/患者ID/区分/来院合計/会計区分/要確認） |
| treatDate 降順ソート + 10件スライス | ✅ 正常 |
| エラーハンドリング | ✅ `{ error: e.message }` 返却 |
| ログ方針（個人情報なし） | ✅ patientId と件数のみ |

---

## 8. selfPayWeb.html（既存稼働導線 — 変更なし確認）

| 確認項目 | 結果 |
|---|---|
| selfPayWeb.html の変更有無 | ✅ 変更なし（WEB-1 では触っていない） |
| `getSelfPayMenuMaster_V3()` | ✅ 存在・変更なし |
| `getSelfPayDataByVisitKey_V3()` | ✅ 存在・変更なし |
| `saveSelfPayDetailsFromDialog_V3()` | ✅ 存在・変更なし |

---

## 9. 総合判定

| ルート | 判定 | 備考 |
|---|---|---|
| `page=search`（デフォルト） | ✅ PASS | 既存稼働導線、コード変更なし |
| `page=selfpay` | ✅ PASS | 既存稼働導線、コード変更なし |
| `page=home` | ✅ PASS（コードレビュー） | 要実機確認 |
| `page=detail&patientId=xxx` | ✅ PASS（コードレビュー） | 要実機確認 |

---

## 10. 要実機確認（現場スマホ）

以下は実機ブラウザでの確認が必要:

1. **`?page=home`** → ナビゲーションカードが表示されるか
2. **`?page=home` → 「患者検索」タップ** → `?page=search` に遷移するか
3. **`?page=search`（デフォルト）** → 既存の検索・選択・自費明細リンク動作するか  
4. **`?page=search` → 患者選択後 → 「患者詳細を見る」タップ** → `?page=detail` に遷移するか
5. **`?page=detail&patientId=実在患者ID`** → 患者情報・来院履歴が表示されるか
6. **`?page=detail` → 「自費明細入力 →」タップ** → `?page=selfpay` に遷移するか

---

## 11. デフォルト URL 変更条件（web-home.html をデフォルト化）

`docs/PHASE_WEB2_VISIT_CREATE_DESIGN_2026-05-05.md §8` 記載の条件リスト:  
1. `web-home.html` 実機確認 PASS ← **未実施**
2. `patientSearch.html` に「← Web ホームへ」リンク追加済み ✅（WEB-1B で実装）
3. `selfPayWeb.html` への既存導線が壊れていない ✅
4. `web-patient-detail.html` 実機確認 PASS ← **未実施**
5. 患者詳細 → 来院記録追加の基本導線が安定（Phase WEB-2 完了後）
6. 現場でスマホ操作が試された実績

→ **条件 1 / 4 / 5 / 6 が未完のためデフォルト変更は保留**
