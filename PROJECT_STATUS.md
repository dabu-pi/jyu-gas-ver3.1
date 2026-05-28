# JREC-01 柔整保険申請書 Ver3.1 — プロジェクトステータス

最終更新: 2026-05-28（**TASK-PORTAL-LINK-AUDIT-003 staff 7 page 戻りリンク追加 + KPI endpoint rollback CLOSED @17/@15**）
担当: dabu-pi
ブランチ: `main`

---

## ✅ 2026-05-28 TASK-PORTAL-LINK-AUDIT-003: 全 staff page に「平山ビジネスポータルへ戻る」追加 + KPI endpoint rollback

状態: **CLOSED** — staff entry @17 / KPI @15（rollback で復旧）

### 実装内容（staff 7 page）

JYU-GAS の全 staff 業務画面に Business Portal 戻りリンク（`target="_blank" rel="noopener noreferrer"`）を追加。

| ファイル | 変更内容 |
|---|---|
| `web-home.html` | Portal-18-D 既存ブロックを `target=_top` → **`target=_blank`** に更新 / rel 強化 / 文言「平山ビジネスポータル」統一 |
| `patientSearch.html` | nav 直前に inline-styled 戻りリンク追加 |
| `selfPayWeb.html` | 同上 |
| `web-monthly-claims.html` | 同上 |
| `web-monthly-claim-detail.html` | 同上 |
| `web-patient-detail.html` | 同上 |
| `web-visit-new.html` | 同上 |

共通 include 機構を JYU-GAS は使っていないため、各 page に individual に inline-styled bar を追加（CSS 重複は inline で回避）。

未関与（staff routing 外）: `web-b2-results.html` / `web-find-months.html` / `web-fixture-results.html`（test/diagnostic 出力）/ `selfPayDialog.html`（Sheets サイドバーダイアログ）

### deploy（staff）

- staff entry: `AKfycbxODNWJ...` @16 → **@17**（URL 不変）

### ⚠️ KPI endpoint 副作用 + Rollback（重要）

JREC AUDIT-002 と同じ dual-deploy ルールで KPI deployment（`AKfycbxNMVV...`）にも @15 → @18 deploy したところ、JBIZ fetchInsuranceKpi が `state: "error"` を返すように regression。

**原因**:
- `?action=insuranceKpiSummary` ハンドラは **このリポに一度も commit されていなかった**（`git log -S"insuranceKpiSummary"` ヒット 0）
- 過去 @14/@15 deploy 時、GAS Apps Script Editor 上で直接コード追加されていた
- 現 HEAD の `Ver3_core.js doGet` には `action` parameter 処理が**完全に無く、すべて `page` routing**
- `clasp push --force` で Editor 上の action handler が **HEAD コードで上書きされ消失**

**rollback**:
- `clasp deploy --deploymentId AKfycbxNMVV... --versionNumber 15` で **KPI deployment を version 15 に戻し復旧**
- 復旧後 JBIZ fetchInsuranceKpi: `{"ok":true, "state":"connected", "data":{insurance_visit_count, ...}}` 正常応答

**結果**:
- staff entry: @17（新 HTML / 戻りリンク強化）
- KPI: **@15 維持**（旧 code / action handler 健在）
- 両 deployment 独立して機能

### ⚠️ JYU-GAS dual-deploy ルール（新規・必須）

> **JYU-GAS の KPI deployment（`AKfycbxNMVV...`）には HEAD code を新規 deploy しない。**
> action handler が GAS Apps Script Editor 上にのみ存在し、HEAD code に commit されていないため、deploy するたびに handler が消滅する。
> 将来 action handler を properly に HEAD へ commit するまで、KPI deployment は **version 15 に pin** する。
> staff entry deployment（`AKfycbxODNWJ...`）への deploy は HEAD code で問題なし。

JREC-SF01 の dual-deploy ルール（staff + KPI 両方 deploy）は JYU-GAS には**適用しない**。

### live 検証

| 項目 | 結果 |
|---|---|
| 7 staff page 全てに戻りリンク表示 + target=_blank | ✅ |
| 患者検索 / 月次申請 / 自費明細 / 患者詳細 / 来院記録 / 月次申請詳細 から戻れる | ✅ |
| KPI endpoint `?action=insuranceKpiSummary` rollback 後 JSON 正常 | ✅（`ok:true / state:connected`）|
| JBIZ Portal-13 fetchInsuranceKpi connected | ✅ |
| JBIZ smoke 262/262 PASS | ✅（regression 0）|

### follow-up（別タスク）

- **TASK-PORTAL-13-INSURANCE-KPI-HANDLER-COMMIT**: 現在 Editor 上にのみ存在する `?action=insuranceKpiSummary` handler を `Ver3_core.js` 等の repo 内 .js ファイルに properly に commit する。完了すれば KPI deployment への HEAD code deploy も安全になり、dual-deploy parity が回復する。

詳細: `hirayama-jyusei-strategy/docs/PORTAL_LINK_AUDIT_2026-05-28.md` §11

---

## 2026-05-14: Git dirty 根本原因解消（緊急対応）

### 発生事象

`gas-projects/jyu-gas-ver3.1` の以下 6 ファイルが HEAD には tracked だが disk から欠損していた:

- `SPEC.md`（20,921 bytes）
- `Ver3_amounts.js`（54,197 bytes）
- `Ver3_core.js`（300,004 bytes）
- `Ver3_patientPicker.js`（6,906 bytes）
- `Ver3_transferData.js`（129,354 bytes）
- `appsscript.json`（244 bytes）

### 危険度

**HIGH** — この状態のまま `clasp push` を実行すると、GAS @13 production code を削除する恐れがあった。

### 対応

1. `git checkout -- <6 files>` で HEAD から復元
2. `git update-index --refresh` + `git status` 空・`git ls-files -d` 空を確認
3. `docs/JYU_GAS_SOURCE_OF_TRUTH_2026-05-14.md` を追加して再発防止ルールを明文化

### 再発防止

- `clasp push` 前に必ず `git ls-files -d` を確認（missing tracked があれば停止）
- ファイル削除は必ず `git rm` + commit まで完結
- workspace 共通 `tools/git-health-check.ps1` を sync 前後で実行

詳細: [`docs/JYU_GAS_SOURCE_OF_TRUTH_2026-05-14.md`](./docs/JYU_GAS_SOURCE_OF_TRUTH_2026-05-14.md)  
workspace root cause: `../../docs/GIT_DIRTY_ROOT_CAUSE_2026-05-14.md`

---

## 現在の状態

**稼働中 + WEB-1〜WEB-4D 実装完了 + デフォルト URL = page=home**

スプレッドシート運用は継続中。  
Web UI から来院記録の登録・候補金額算定まで実装済み（needCheck=TRUE / 要確認）。  
`/exec` のデフォルトを `page=home`（ナビゲーションハブ）に変更済み。  
`/exec?page=search` で患者検索、`/exec?page=selfpay` で自費明細は従来通り。

### デフォルト URL 変更（2026-05-06）

| 変更 | 内容 |
|---|---|
| 変更前 | `/exec` → patientSearch.html |
| 変更後 | `/exec` → web-home.html |
| page=search | `/exec?page=search` で従来通りアクセス可能 |
| 既存導線 | patientSearch / selfPayWeb は変更なし |
| deployment | @9 → @10 に更新（同一 deploymentId） |

**重要: clasp push だけでは /exec は更新されない**

### 白画面バグ修正（2026-05-06）

**原因:** GAS Web App の 2 段 iframe 構造内で `<a href>` をクリックすると、
トップウィンドウではなく内側 iframe 内で遷移してしまい、GAS 構造が入れ子になる。

**修正:** 全ページ遷移リンクに `target="_top"` を追加。
- `patientSearch.html`: 「← Web ホームへ」「患者詳細を見る」「自費明細入力」
- `web-home.html`: 「患者検索」カード
- `web-patient-detail.html`: 「患者検索に戻る」「来院記録を追加」「自費明細入力」
- `web-visit-new.html`: 「患者詳細に戻る」（静的 + JS 生成）

**本番反映:** deployment @10 → @11 （clasp deploy -i）  
`clasp push` は HEAD のみ更新。`/exec` の反映には `clasp deploy -i <deploymentId>` が必要。

**本番確認コマンド:**
```powershell
npx tsx tools/live-check-runner/scripts/check-exec-home.ts
```

### 次のアクション

**→ WEB-6 本番 deploy 完了（2026-05-08）** ★ — 共通グローバルナビタブ @13 反映  
  - 本番確認済み: 全6ページ nav / target="_top" / 白画面なし / Step1 金額 ¥4,363/¥3,053

**→ WEB-6 実装完了（2026-05-08）** — 共通グローバルナビタブ追加  
  - 7ページに `.web-nav` タブナビを追加（全リンク `target="_top"` 付き）
  - 対象: web-home / patientSearch / web-patient-detail / web-visit-new / web-monthly-claims / web-monthly-claim-detail / selfPayWeb  
  - 旧ホームリンクを削除（patientSearch・web-monthly-claims）  
  - clasp push 済 / **dev 目視確認待ち** / 本番 deploy 未実施

**→ WEB-5 本番フロー一周確認完了（2026-05-08）** ★  
  - 本番 @12 全フロー確認済み（自動スクリプト）  
  - 月次一覧: 9名 / hirayamaka 確認  
  - Step1後: ¥4,363 / ¥1,310 / ¥3,053（転記データ）✅  
  - B案実生成: `申請書_hirayamaka_2026-04_151333.xlsx` ✅  
  - APPGEN_SECRET 露出なし ✅ / A案PDF未使用 ✅  
  - 残課題: 一覧側金額は来院ヘッダ値（5円差）→ 次フェーズ候補

**→ WEB-4A〜4D 本番 deploy 完了（2026-05-08）** ★  
  - deployment: @12 / deploymentId: `AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_Ilt8SS5P5zodfF2dnmKeqso8BL8hcinVEBrQ`  
  - 本番 `/exec` 確認済み:  
    - `?page=monthlyClaims`: ✅ 表示  
    - `?page=monthlyClaimDetail`: ✅ 表示  
    - Step1後 来院合計 ¥4,363 / 窓口 ¥1,310 / 請求 ¥3,053 ✅  
    - tfoot・KPI・注記すべて更新 ✅  
    - B案Excel生成ボタン表示 ✅  
    - APPGEN_SECRET 露出なし ✅

**→ WEB-4C 修正完了（2026-05-08）** — Web月次集計とB案申請書の金額整合  
  - 原因: 来院ヘッダの per-visit 候補金額合算 vs `V3TR_buildTransferDataForMonth_` 月合計再計算の丸め差（5円）  
  - 修正: Step1成功後のハンドラで KPI を転記データ金額で上書き + 注記表示  
  - 修正後: Step1後 KPI ¥4,363 / ¥1,310 / ¥3,053 = B案 Excel と一致  
  - LiveCheck W4C-1〜5 全 PASS / B案ルート変更なし / clasp push 済

**→ WEB-4B 修正完了（2026-05-08）** — 月次申請集計0円バグ修正  
  - 原因: `getMonthlyClaimList_V3` / `getMonthlyClaimDetail_V3` で `buildHeaderColMap_`（1始まり）を配列インデックス（0始まり）として使用 → 列ズレ → 全行スキップ → 集計0  
  - 修正: 両関数を `V3TR_buildHeaderMap_`（0始まり）に変更  
  - 修正後: 来院数2 / ¥4,368 / ¥1,310 / ¥3,058 を正しく表示  
  - LiveCheck W4B-1〜5 全 PASS / B案ルート変更なし / clasp push 済

**→ WEB-4A 実機確認完了（2026-05-08）** — Web UI から B案 Cloud Run Excel 申請書生成入口を確認  
  - `generateClaimApplicationBFromWeb_V3(patientId, ym)` 追加（Ver3_transferData.js）
  - `web-monthly-claim-detail.html` に Step 3「申請書Excelを生成」ボタン追加
  - LiveCheck W4A-1〜5 全 PASS（auth更新後）
  - 実生成確認: `申請書_hirayamaka_2026-04_102943.xlsx` 生成・Drive保存・ログ記録 ✅
  - 回帰: smoke 28 PASS / web3 8 PASS / web34 9 PASS / 1 SKIP（設計通り）
  - clasp push 済 / **本番 deploy 未実施**
  - 残人間確認: 印刷プレビュー1ページ・転帰欄丸囲み・罫線位置（Drive URL はログシートで確認）

**→ WEB-3.4 実装完了・LiveCheck 9 PASS / 1 SKIP（2026-05-07）**  

**→ テストデータ削除 COMPLETED（2026-05-07）**  
  - `hirayamaka_2998-12-31`（来院ケース・来院ヘッダ・施術明細）— 削除済み  
  - `hirayamaka_2999-12-31`（来院ケース・来院ヘッダ）— 削除済み  
  - devExecuteCleanupTestVisitData_V3 再実行で「合計: 0 行」を確認

**→ B-1 fixture テスト COMPLETED（2026-05-07）**  
  - TC01〜TC25b 全57ケース PASS（LiveCheck: npm run test:jyu:fixtures）

**→ B-2 インフラ検証 COMPLETED（2026-05-07）**  
  - verifyMonthlyClaimData_V3 正常動作 / 6 PASS  
  - 当月（2026-05）: NO_PATIENTS_THIS_MONTH（実来院データなし）  
  - 実データ検証は月次来院後に人間が実施

**→ web251 SKIP → 4 PASS 解消確認（2026-05-07）**  
  - テストデータ削除後 1 回目実行: 4 PASS ✅  
  - 2 回目以降はデータ再作成で SKIP（設計通り）  
  - 完全 4 PASS 化には `devCleanupTestVisitData_V3(false)` の都度実行が必要

**【前フェーズ 残人間タスク（アーカイブ）】**  
1. ~~`?page=b2Results&ym=YYYY-MM` で実来院月を指定して B-2 実データ確認~~ → COMPLETED  
2. ~~申請書テンプレート（「新 様式第5号」シート）の有無をスプレッドシートで確認~~ → COMPLETED  
3. 実患者・実月で Excel 目視確認（上記残人間タスク §1 に移行）  
4. 現場スマホ実機確認（チェックリスト: `docs/WEB25_SMARTPHONE_FIELD_CHECK_2026-05-06.md`）  
5. WEB-3.4 本番 deploy（月次確認後）

**→ B-3 COMPLETED（2026-05-07）** — SPEC.md 新規作成・§14 に Web 登録フロー仕様追記  
**→ auth 更新後 回帰テスト CONFIRMED（2026-05-07）** — 61 PASS / 2 SKIP / 0 FAIL  
**→ B-2 実データ確認 COMPLETED（2026-05-07）** — PDF生成成功・新様式第5号確認  
**→ 申請書出力方式確定 + ツール整備完了（2026-05-07）**  
  - **推奨: NDJSON + Python スクリプト（write_application.py）**  
    → tools/claim-excel/ に write_application.py + application_template.xlsx を配置済み  
    → verify_application_xlsx.py（自動確認スクリプト）も追加  
    → GAS 側: exportClaimNdjson_V3 + Web UI ボタン実装済み  
  - **試験的: Sheets直PDF（generateClaimApplication_V3）**  
    → 施術日カレンダー ○・転帰が未実装のため帳票として不完全  
  - **DEPLOY 保留: Excel 目視確認後に判断**  
**→ 申請書生成B案 正ルート採用確定 COMPLETED（2026-05-07）** ★
  - 既存メニュー「帳票出力 → 申請書を出力」= `V3TR_menuGenerateApplication_B()` (`Ver3_transferData.js`)
  - Cloud Run `jrec-appgen-server` 稼働確認 / URL: `https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app`
  - `/health` → `{"status":"ok"}` ✅ (revision 00026-wv2 / 2026-04-20 デプロイ)
  - B案実行 (hirayamaka/2026-04): `申請書_hirayamaka_2026-04.xlsx` 生成 (36,694 bytes) ✅
  - **スプレッドシートメニューから実行・Drive保存・申請書生成ログ記録 — 全て正常動作**
  - **人間目視確認済み（2026-05-07）:**
    - 負傷名: （1）頸部 捻挫 / （2）腰部 捻挫 — 正しく連続・（3）以降に残存なし ✅
    - 施術日: ① ⑲（4/1・4/19）が入力されている ✅
    - 合計 4,363円 / 一部負担金 1,310円 / 請求金額 3,053円 ✅
    - 申請書生成ログにも本日実行分が記録されている ✅
  - 生成ファイル Drive URL: https://docs.google.com/spreadsheets/d/1HHv-KH6bmfA-vEZErZ94uZxLutsxGi-C/edit
  - Sheets直PDF (A案) は停止状態を維持（施術日カレンダー○・転帰が未実装）

**→ NDJSON+Python ローカル実行 COMPLETED（2026-05-07）**
  - `tools/claim-excel/write_application.py` でも同一 Excel 生成可能（補助ルート）
  - `verify_application_xlsx.py` 実行 → E26 = `（1）頸部 捻挫` ✅

---

## 申請書出力ルート確定（2026-05-07）★ 採用確定

| ルート | 位置づけ |
|---|---|
| **B案（Cloud Run → xlsx → Drive）** | ✅ **正ルート・採用確定** |
| Sheets直PDF（A案） | ❌ 停止・本番化しない |
| NDJSON+Python ローカル | 補助（B案と同一ロジック / 開発時使用可） |

**B案 = スプレッドシートメニュー「帳票出力 → 申請書を出力」**

- GAS関数: `V3TR_menuGenerateApplication_B()` (`Ver3_transferData.js`)
- Cloud Run: `https://jrec-appgen-server-j6vlxdvqaa-an.a.run.app`
- GASスクリプトプロパティ: `APPGEN_ENDPOINT` + `APPGEN_SECRET` 設定済み（運用中）

**Web UI からB案を呼ぶ将来構成（次回以降の候補）:**
```
Web → google.script.run → GAS wrapper → V3TR_generateApplicationBCore_()
→ Cloud Run → Drive保存 → WebへURL返却
```
APPGEN_SECRET を Web/JS に出さないこと。既存B案ロジックを壊さないこと。

**【残人間確認（目視・印刷）】**
- [ ] 印刷プレビューで 1 ページに収まるか
- [ ] 転帰欄の丸囲みの見た目
- [ ] 罫線・文字位置の最終確認

**本番 deploy: 保留**（B案は Cloud Run xlsx のため GAS deploy 対象外）

**本日（2026-05-07）はここで作業終了。新規実装・clasp deploy・本番反映は行わない。**

---

## 実装フェーズ一覧

| フェーズ | 内容 | 状態 |
|---|---|---|
| WEB-1 | Web UI 入口・患者詳細・home 画面 | ✅ 完了 |
| WEB-2 | Web UI 来院登録（金額=0・要確認） | ✅ 完了 |
| WEB-2.5 | Web UI 来院登録 × 候補金額算定 | ✅ 完了 |
| スマホ実機確認 | 現場スマホでの動作確認 | ✅ Playwright mobile PASS / 実機確認待ち |
| WEB-2.5.1 | 施術明細自動生成 | ✅ CLOSED（LiveCheck 4 PASS / 2026-05-07） |
| WEB-3 | 月次申請フロー（一覧・詳細・転記データ生成） | ✅ 完了（LiveCheck 8 PASS / 2026-05-07） |
| WEB-3.4 | 申請書 PDF 生成（A案：テンプレ書込 + Drive PDF） | ✅ 完了（LiveCheck 9 PASS / 1 SKIP / 2026-05-07） |
| WEB-4A | Web UI から B案 Cloud Run Excel 申請書生成入口 | ✅ 完了（clasp push 済 / 2026-05-08） |
| WEB-4B | 月次申請集計0円バグ修正（buildHeaderColMap_ off-by-one） | ✅ 完了（clasp push 済 / 2026-05-08） |
| WEB-4C | Web月次集計とB案申請書の金額整合（Step1後KPI上書き） | ✅ 完了（clasp push 済 / 2026-05-08） |
| WEB-4D | tfoot合計行更新バグ修正（cells.length >= 6 → >= 4） | ✅ 完了（clasp push 済 / 2026-05-08） |
| **本番 deploy** | WEB-4A〜4D 本番反映 @12 | ✅ **2026-05-08 完了** |
| WEB-6 | 共通グローバルナビタブ追加（7ページ） | ✅ **本番 deploy @13 完了（2026-05-08）** |

---

## Phase WEB-1 実装内容（2026-05-05）

### 新規追加ファイル

| ファイル | 内容 |
|---|---|
| `web-home.html` | Web ナビゲーションハブ（カード形式） |
| `web-patient-detail.html` | 患者詳細・来院履歴の読み取り専用画面 |
| `docs/WEB_UI_MIGRATION_PLAN_2026-05-05.md` | Web UI 移行計画 Markdown |

### Ver3_core.js 変更内容

| 変更種別 | 内容 |
|---|---|
| `doGet(e)` 拡張 | `page=home` / `page=detail` ルート追加（既存ルートは変更なし） |
| 新関数追加 | `getPatientDetail_V3(patientId)` — 患者基本情報 + 来院履歴10件返却 |

### patientSearch.html 変更内容

- 「患者詳細を見る」ボタンを選択後パネルに追加（`?page=detail&patientId=xxx` へのリンク）

---

## doGet ルーティング（現在）

| `page=` | HTML | 状態 |
|---|---|---|
| `search`（デフォルト） | `patientSearch.html` | 稼働中 |
| `selfpay` | `selfPayWeb.html` | 稼働中 |
| `home` | `web-home.html` | WEB-1 追加 |
| `detail` | `web-patient-detail.html` | WEB-1 追加 |

---

## 既存スプレッドシート運用への影響

**影響なし**

- 既存シート構造は変更していない
- 既存 GAS 関数は変更していない（doGet の既存ルートも変更なし）
- スプレッドシート操作（保存・帳票出力）は従来通り動作する

---

## Phase WEB-1B 入口整理・設計固め（2026-05-05）

### 実施内容

| 項目 | 内容 |
|---|---|
| `patientSearch.html` | 「← Web ホームへ」リンクを追加（最小差分） |
| デフォルト URL 方針 | `page=search` のまま維持（変更しない） |
| 設計 Markdown 作成 | `docs/PHASE_WEB2_VISIT_CREATE_DESIGN_2026-05-05.md` |

### デフォルト URL をまだ変更しない理由

`patientSearch.html` → `selfPayWeb.html` はスマホ実地テスト済みの稼働導線。
`web-home.html` は実機未確認のため、デフォルトに変更すると既存運用が止まるリスクがある。
条件リスト（PHASE_WEB2 設計 §8）が揃ったタイミングで変更する。

### `saveVisitFromWeb_V3` が必要な理由

既存の `saveVisit_V3` はスプレッドシートの患者画面シート（C2, B4, A12:H13 等）からデータを読むため、
Web App の google.script.run から呼んでも意図した値が読めない。
JSON 引数で受け取る専用関数 `saveVisitFromWeb_V3(payload)` を新規実装する必要がある。

### `setPatientAndDate_V3` を Web 保存処理に流用しない理由

B2（患者表示名）と B4（来院日）を **シートに書き込む副作用** がある。
Web セッションと並行してスプレッドシートを開いている場合に干渉する。
Web UI からの来院登録では、シートを経由しない保存経路（`saveVisitFromWeb_V3`）を使う。

### 次の実装候補（Phase WEB-2）

```
1. getPrevVisitData_V3(patientId)  ← 前回来院データ JSON 返却
2. web-visit-new.html              ← 来院登録フォーム
3. doGet に page=visitNew 追加
4. saveVisitFromWeb_V3(payload)    ← 来院登録保存（UIシート非依存）
5. web-patient-detail.html に「来院記録を追加」ボタン
```

---

## Phase WEB-1 後 既存 Web UI 棚卸し（2026-05-05）

詳細は `docs/WEB_UI_EXISTING_INVENTORY_2026-05-05.md` 参照。

### 棚卸し結果サマリ

- 既存稼働中導線: `patientSearch.html` → `selfPayWeb.html`（スマホ操作・実地テスト済み）
- Phase WEB-1 の位置づけ: **部分延長 + 部分並列**
  - 延長: `patientSearch.html` に「患者詳細を見る」追加
  - 並列: `web-home.html` は `?page=home` 指定でのみアクセス可能（デフォルト URL は変更なし）
- 既存スプレッドシートUI専用関数（Web から直接呼べない）: `saveVisit_V3`, `autofillFromPreviousVisit_V3`, `openSelfPayDialog_V3`

### Phase WEB-2 前に決めること

1. **デフォルト URL の方針**（`web-home.html` をデフォルトにするか）
2. **来院登録の設計**（`saveVisitFromWeb_V3` 新関数の引数・バリデーション方針）
3. **実機確認**（`web-home.html` / `web-patient-detail.html` の動作確認）

### Phase WEB-2 で必要な新関数

| 関数 | 役割 |
|---|---|
| `saveVisitFromWeb_V3(params)` | JSON 引数で来院登録（`saveVisit_V3` の Web 版） |
| `getPrevVisitData_V3(patientId, treatDate)` | 前回来院データの JSON 返却 |

---

## Phase WEB-2.5 調査・設計（2026-05-06）

**ステータス: 実装完了 / LiveCheck 5 PASS**

設計書: `docs/WEB25_AMOUNT_CALCULATION_DESIGN_2026-05-06.md`

### 調査結論

| 確認事項 | 結論 |
|---|---|
| `calcHeaderAmountsByVisitKey_V3_` を Web から呼べるか | ✅ 全シート読み取りのみ・UIシート依存なし |
| `calcEpisodeForCase_` を Web から呼べるか | ✅ 来院ケースのみ読み取り・UIシート依存なし |
| `saveVisit_V3` のロジックを流用できるか | △ UIシート部分を除外し、算定関数のみ再利用 |
| kubun を自動判定できるか | ✅ `calcEpisodeForCase_` で30日ルール準拠 |
| 施術明細・初検情報履歴の書き込みが必要か | △ MVP では省略可能（needCheckReason で記録） |

### WEB-2.5 実装内容（2026-05-06）

| 実装 | 結果 |
|---|---|
| `saveVisitFromWeb_V3` 改修: kubun を `calcEpisodeForCase_` で自動判定 | ✅ 完了 |
| `saveVisitFromWeb_V3` 改修: `calcHeaderAmountsByVisitKey_V3_` で候補金額算定 | ✅ 完了 |
| 来院ヘッダに候補金額保存（needCheck=true + 理由付き） | ✅ 完了 |
| `web-visit-new.html` 成功画面に候補金額表示 | ✅ 完了 |
| LiveCheck（W2.5-1〜5 全 PASS） | ✅ 完了 |
| clasp push | ✅ 完了 |

**LiveCheck テスト visitKey:** `hirayamaka_2999-12-31`

### テストデータ確認結果（2026-05-06）

| 確認項目 | 結果 |
|---|---|
| 来院ヘッダ存在 | ✅ 確認済み |
| 来院合計 ¥2,410 | ✅ 確認済み |
| needCheck=TRUE | ✅ 確認済み |
| 区分 = 初検（自動判定） | ✅ 確認済み |
| 削除 | ✅ 削除確認済み（W2.5-4 新規保存 PASS で確認） |

**削除確認済み:** W2.5-4 が「新規保存 PASS」（✅）になったことで、  
元のテストデータが正常に削除されていたことを確認。  
削除後の保存で来院合計 ¥2,410・区分 初検（自動判定）が正常動作。

W2.5-4 の実行により `(検証用実在ID)_2999-12-31` が再作成されています。  
引き続きスプレッドシートで削除可能（施術日 2999-12-31 の行）。

**WEB-2.5 LiveCheck 全テスト: 5 passed / 0 failed**
- W2.5-1: kubun 未選択でモーダルが開く ✅
- W2.5-2: 「システムが自動判定」モーダル表示 ✅
- W2.5-3: 「請求確定ではありません」警告表示 ✅
- W2.5-4: 新規保存 PASS（削除確認） ✅
- W2.5-5: 二重保存防止 DUPLICATE_VISIT ✅

### 実装後の動作仕様

- kubun は `calcEpisodeForCase_`（30日ルール）で自動判定（user input は参考のみ）
- 来院ヘッダに候補金額（initFee / reFee / visitTotal / windowPay / claimPay）を記録
- `needCheck=true` は常に維持
- `needCheckReason` = "Web UI 登録;（算定抑制理由）;施術明細未記録（Web MVP）;（初検時: 初検情報履歴未記録）"
- 成功画面: visitKey / 区分（自動判定）/ 来院合計 / 窓口負担 / 保険請求 / 要確認理由を表示

### 実装前オーナー確認事項

| 確認 | 内容 |
|---|---|
| A | Web 登録の needCheck=true を月次申請前に必ず確認する運用を合意 |
| B | 施術明細なしヘッダが既存処理に影響しないこと |
| C | Web 登録の kubun を後変更する手順（`saveVisit_V3` で上書き可） |
| D | W2.5 テスト実施時のデータ削除手順 |

---

## Phase WEB-2 実装内容（2026-05-06）

### 追加ファイル

| ファイル | 内容 |
|---|---|
| `web-visit-new.html` | 来院記録登録フォーム（WEB-2 新規） |
| `docs/WEB1_LIVECHECK_2026-05-06.md` | WEB-1 コードレビュー LiveCheck 記録 |
| `docs/WEB2_LIVECHECK_2026-05-06.md` | WEB-2 コードレビュー LiveCheck 記録 |

### Ver3_core.js 追加関数

| 関数 | 行番号 | 役割 |
|---|---|---|
| `getPrevVisitData_V3(patientId)` | 5831 | 前回来院データを JSON 返却（読み取りのみ） |
| `findLatestCaseKeyForPatient_(caseSh, caseMap, pid, caseNo)` | 5924 | 最新 caseKey 取得ヘルパー |
| `saveVisitFromWeb_V3(payload)` | 5976 | 来院を来院ケース+来院ヘッダに保存（UI シート非依存） |

### doGet 変更

| page= | HTML | 状態 |
|---|---|---|
| `visitNew` | `web-visit-new.html` | WEB-2 追加 |

### web-patient-detail.html 変更

- 「来院記録を追加 →」ボタンを追加（`?page=visitNew&patientId=xxx` へ遷移）
- 「自費明細入力 →」ボタンは変更なし

### WEB-2 の制約（意図的・スコープ外）

| 制約 | 理由 |
|---|---|
| 金額計算なし（来院合計=0, 要確認=true） | calcVisitAmounts_V3_ はシート依存が深い（WEB-2.5 予定） |
| ケース1のみ（ケース2は未対応） | MVP に不要（WEB-2.5 予定） |
| keikaNow / shoken 入力なし | セル結合の複雑さ（WEB-2.5 予定） |
| 保険算定なし | スプレッドシートで従来通り実施 |

---

## 現在の doGet ルーティング

| `page=` | HTML | 状態 |
|---|---|---|
| `search`（デフォルト） | `patientSearch.html` | 稼働中 |
| `selfpay` | `selfPayWeb.html` | 稼働中 |
| `home` | `web-home.html` | WEB-1 追加（実機未確認） |
| `detail` | `web-patient-detail.html` | WEB-1 追加（実機未確認） |
| `visitNew` | `web-visit-new.html` | WEB-2 追加（実機未確認） |

---

## Playwright LiveCheck 状況（2026-05-06）

**スペック:** `tools/live-check-runner/projects/jyu-gas-ver31/`

| `web3.spec.ts`（WEB-3 W3-1〜8） | `npm run test:jyu:web3` | **8 PASS / 0 FAIL / 0 SKIP** | ✅ CLOSED 2026-05-07 |
| `web34.spec.ts`（WEB-3.4 W3.4-1〜10） | `npm run test:jyu:web34` | **9 PASS / 1 SKIP / 0 FAIL** | ✅ CLOSED 2026-05-07 |
| `tc_fixtures.spec.ts`（B-1 TC-ALL/DETAIL） | `npm run test:jyu:fixtures` | **2 PASS / 0 FAIL** (57 fixture PASS) | ✅ CLOSED 2026-05-07 |
| `b2_transfer.spec.ts`（B-2 B2-1〜10） | `npm run test:jyu:b2` | **6 PASS / 0 FAIL / 0 SKIP** | ✅ インフラ検証完了 2026-05-07 |
| `b2_realdata.spec.ts`（B-2 実データ B2R-1〜7） | `npm run test:jyu:b2real` | **2 PASS / 0 FAIL / 0 SKIP** | ✅ 実データ確認完了（2026-04 / PDF生成成功）|
| `find_months.spec.ts`（月スキャン） | `npm run test:jyu:findmonths` | **1 PASS** | ✅ 2026-04 = 9患者 検出 |

### 回帰テスト合計（2026-05-07 最新）

| suite | 結果 | 備考 |
|---|---|---|
| smoke | 28 PASS | |
| web25 | 5 PASS | |
| web251 | 3 PASS / 1 SKIP | W2.5.1-1 がテストデータ再生成 → 2回目以降設計通り SKIP |
| web3 | 8 PASS | |
| web34 | 9 PASS / 1 SKIP | W3.4-10 inner frame evaluate SKIP（設計通り） |
| fixtures (B-1) | 2 PASS (57 fixture PASS) | TC01〜TC25b 全件 PASS |
| b2 transfer | 6 PASS | NO_PATIENTS_THIS_MONTH（2026-05 データなし） |
| findmonths | 1 PASS | 2026-04 / 9患者 |
| b2real | 2 PASS (PDF成功) | 2026-04 / ¥3,053 / 84セル |
| **合計（B-2 実データ確認後）** | **63 PASS / 2 SKIP / 0 FAIL** | 2026-05-07 確認済み |

---

## Phase WEB-3 実装内容（2026-05-07）

**ステータス: 完了（LiveCheck 8 PASS）**

設計書: `docs/WEB3_MONTHLY_CLAIMS_DESIGN_2026-05-07.md`

### 新規 GAS 関数

| 関数 | 役割 |
|---|---|
| `getMonthlyClaimList_V3(ym)` | 月次申請対象者一覧（既存 V3TR_findPatientsForMonth_ 再利用） |
| `getMonthlyClaimDetail_V3(patientId, ym)` | 患者×月 来院詳細（読み取り専用） |
| `buildMonthlyTransferData_V3(patientId, ym)` | 転記データ生成（既存 V3TR_buildTransferDataForMonth_ ラップ） |

### 新規 HTML ページ

| ファイル | page= | 内容 |
|---|---|---|
| `web-monthly-claims.html` | `monthlyClaims` | 月次申請対象者一覧（年月入力 → テーブル） |
| `web-monthly-claim-detail.html` | `monthlyClaimDetail` | 詳細・プレビュー・転記データ生成 |

### doGet 変更

| 追加ルート | HTML |
|---|---|
| `page=monthlyClaims` | web-monthly-claims.html |
| `page=monthlyClaimDetail` | web-monthly-claim-detail.html |

### web-home.html 更新

- 「来院記録」カードを disabled → active（?page=visitNew リンク）
- 「月次申請」カードを新規追加（active / ?page=monthlyClaims リンク）

### 制度準拠

- 算定事実のみ表示（来院ヘッダの確定値）
- 自費のみ来院を除外（会計区分フィルタ）
- needCheck=true は「要確認」バッジで明示
- 請求確定は Sheets UI で人間が行う運用を維持
- 申請書 PDF 生成は既存メニュー経由（WEB-3.4 で Web 対応予定）

### Dashboard 関連

なし（対象外）

---

## Playwright LiveCheck 状況（最新版）

**スペック:** `tools/live-check-runner/projects/jyu-gas-ver31/`

| スペック | コマンド | **最終結果** | 備考 |
|---|---|---|---|
| `smoke.spec.ts`（WEB-1 S-1〜S-6） | `npm run test:jyu:smoke` | **26 PASS / 0 FAIL** | 全件 PASS（2026-05-06） |
| `web2.spec.ts`（WEB-2 W2-1〜W2-8） | `npm run test:jyu:web2` | **16 PASS / 0 FAIL** | 全件 PASS（2026-05-06） |
| **合計（〜2026-05-06）** | `npm run test:jyu` | **42 PASS / 0 FAIL / 0 SKIP** | ✅ 完了 |
| `web251.spec.ts`（WEB-2.5.1 W2.5.1-1〜4） | `npm run test:jyu:web251` | **4 PASS / 0 FAIL / 0 SKIP** | ✅ CLOSED 2026-05-07 |

### テスト修正一覧（2026-05-06）

| テスト | 原因 | 修正 |
|---|---|---|
| S-3 | `isVisible()` リトライなし → inner iframe 描画前に false | `waitFor({ state: "visible" })` に変更 |
| S-6 | `#loading` が DOM 未ロード時 → `waitFor(hidden)` が即時解決 | `Promise.race` 両腕を `waitFor({ state: "visible" })` に変更 |
| S-6/W2 | `testData.patientId` 未設定 | `find-patient-id.ts` で検証用実在IDを取得・設定 |

**GAS コード変更なし。doGet のデフォルト（page=search）は正常維持。**

### auth 更新手順（次回確認時）

```powershell
# 1. Chrome を remote debugging で起動
$dir = "C:\hirayama-ai-workspace\workspace\tools\live-check-runner\.chrome-profile"
Start-Process "chrome" "--remote-debugging-port=9222 --user-data-dir=`"$dir`""

# 2. Chrome で以下を開く（順番通りに実行）
#    a) https://accounts.google.com でログイン確認（RTS 更新のため必須）
#    b) JYU-GAS dev URL を開き Account Chooser で pinshanka24@gmail.com を選択
#    c) GAS ページが表示されるまで待つ（この確認が重要）

# 3. GAS ページ表示確認後に save-auth
cd C:\hirayama-ai-workspace\workspace\tools\live-check-runner
npm run save-auth

# 4. テスト実行
npm run test:jyu
npm run test:jrec:smoke
```

### testData.patientId が必要なテスト

`projects/jyu-gas-ver31/config.json` の `testData.patientId` に実在 ID を設定後:
- W2-3b, W2-4b, W2-6/7, S-6a/b が PASS になる

---

## 要実機確認（現場スマホ）

**WEB-1:**
1. `?page=home` → ナビゲーション表示
2. `?page=search` → 検索・選択・自費明細リンク動作
3. `?page=detail&patientId=実在ID` → 患者情報・来院履歴表示
4. `?page=detail` → 「来院記録を追加」「自費明細入力」ボタン動作

**WEB-2:**
1. `?page=visitNew&patientId=実在ID` → フォーム表示
2. 「前回引き継ぎ」→ 前回データがセットされる
3. フォーム入力 → 確認モーダル → 登録実行
4. 来院ケースシート・来院ヘッダシートに行が追加されることを確認
5. 来院ヘッダの 要確認=TRUE / 来院合計=0 を確認

---

---

## Phase WEB-2.5.1 実装内容（2026-05-07）

**ステータス: CLOSED（LiveCheck 4 PASS / 2026-05-07）**

### 概要

`saveVisitFromWeb_V3` に施術明細シートへの自動 upsert を追加した。
WEB-2.5 で算定済みの `amounts.details`（部位別明細）を、既存の `upsertDetailRows_V3_` 関数を使って施術明細シートに書き込む。

### 変更ファイル

| ファイル | 変更内容 |
|---|---|
| `Ver3_core.js` | `saveVisitFromWeb_V3` に 3 箇所変更（最小差分） |

### 変更内容（Ver3_core.js）

| 変更種別 | 内容 |
|---|---|
| 変数追加 | `var epByNo = {}` — caseNo→ep オブジェクトの保持 |
| ループ内追加 | `epByNo[caseNo] = ep` — 施術明細 upsert に必要な episodeStartDate を保持 |
| 理由文変更 | `"施術明細未記録（Web MVP）"` → `willWriteDetails が false のときのみ "施術明細未記録"` |
| ステップ④.5追加 | `upsertDetailRows_V3_` 呼び出し（保険来院かつ amounts.details が存在する場合） |

### 設計上の重要点

| 項目 | 内容 |
|---|---|
| 保険来院のみ | `isInsuranceVisit && amounts.details` の場合のみ施術明細を書く |
| 自費のみ来院 | `buildZeroInsuranceAmounts_V3_()` は details なし → 施術明細未記録が needCheckReason に残る |
| ep1/ep2 の保証 | ケースが1つのみの場合でも fallback `{ episodeStartDate: visitDate }` で安全 |
| needCheck=true | 施術明細を書いても needCheck は常に true（月次確認必須の原則を維持） |
| `upsertDetailRows_V3_` の再利用 | saveVisit_V3（スプレッドシートUI）と完全に同じ関数を使用 |

### clasp push

完了（2026-05-07 08:38:48）

### LiveCheck

| コマンド | 結果 | 実施日 |
|---|---|---|
| `npm run test:jyu:web251` | **4 PASS / 0 FAIL / 0 SKIP** | 2026-05-07 |

**W2.5.1-1 保存結果（実測値）:**
```
visitKey:   hirayamaka_2998-12-31
区分:       初検（自動判定）
来院合計:   ¥2,410
窓口負担:   ¥720
保険請求:   ¥1,690
要確認理由: Web UI 登録;温罨法 算定不可（初検日特例：捻挫）;初検情報履歴未記録（Web MVP）
```
「施術明細未記録」が要確認理由に含まれていない → WEB-2.5.1 の核心が正しく動作している。

**W2.5.1-2〜4:** W2.5.1-1 で保存済み（DUPLICATE_VISIT = 正常動作として PASS）

**テストデータ（要手動削除）:**
- `hirayamaka_2998-12-31` — 来院ケース・来院ヘッダ・施術明細シートの該当行を削除
- `hirayamaka_2999-12-31` — WEB-2.5 テストデータが残存。同様に削除

**施術明細シート確認（手動）:**
W2.5.1-1 の保存により `hirayamaka_2998-12-31` の施術明細行が書き込まれているはず。
スプレッドシートの「施術明細」シートで visitKey 列を確認すること。

### 施術明細シートへの書き込み手動確認手順

1. GAS dev URL の visitNew フォームから実際に来院登録
2. スプレッドシートの「施術明細」シートを開く
3. 該当 visitKey の行が追加されていることを確認
4. 各列（bui / byomei / kubun / 金額列）が正しく入っていることを確認

---

## 次フェーズ候補

### Phase WEB-2.5
Web 登録後の金額算定と拡張。

- `saveVisitFromWeb_V3` から `calcVisitAmounts_V3_` を呼び出す
- ケース2入力対応
- keikaNow / shoken 入力フォーム

### Phase WEB-3
Web から施術録・申請書生成へ。

- 月次申請対象者一覧
- 申請書プレビュー・生成
- PDF / 印刷導線

---

---

## 設計方針リファレンス

詳細は `docs/WEB_UI_MIGRATION_PLAN_2026-05-05.md` を参照。

### 個人情報ログ禁止フィールド

氏名・住所・電話番号・生年月日・保険者番号・記号番号・被保険者情報

### 算定ルール優先順位（維持）

30日ルール → 月内制御 → 区分確定 → 逓減 → 長期減額
