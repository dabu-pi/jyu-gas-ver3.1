# JREC-01 次フェーズ実装方針 — 2026-06-03（設計フェーズ）

本ファイルは「JREC-01 保険対応アプリを、制度遵守・監査耐性・事故ゼロ設計に近づける」次フェーズを**1つに絞って確定**するための設計記録。
このフェーズではコード実装・GAS push・Cloud Run deploy・本番書込を一切行っていない（docs のみ）。

前提（2026-06-03 時点）:
- repo: `main` / HEAD `171f986` / clean / 0-0 / missing 0
- 本体SS: 【毎日記録】来店管理施術録ver3.1 / `1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM`
- 申請書正ルート: B案 Cloud Run Excel（`V3TR_menuGenerateApplication_B`）
- deployment: staff @17 / KPI @15(pin)

---

## 1. 次フェーズ候補の比較

評価軸: 返戻・査定リスク低減 / 制度遵守 / 本番書込事故回避 / live-check 検証容易性 / 外販耐性 / 監査ログ残存

### A. 保険対応アプリ UI/導線整理

```
目的:        Web UI のナビ・導線の使い勝手向上
対象ファイル: web-*.html
対象関数:    doGet ルーティング
対象URL:     staff /exec @17
対象SS:      JREC本体（読み取り中心）
本番書込リスク: 低（UI のみ）
GAS deploy 必要性: あり（HTML変更は clasp deploy -i 必須）
Cloud Run deploy: 不要
live-check 確認可否: 可（要 auth 更新）
人の確認作業: 現場スマホ目視
制度リスク:   低（制度ロジック非該当）
優先度:      低
先にやる理由: 特になし（制度・監査に寄与しない）
後回し可:    UI は WEB-6 まで整備済み。制度遵守の本丸ではない
```

### B. live-check-runner auth 更新 + read-only smoke

```
目的:        Playwright 検証基盤の復旧（auth.json 失効解消）
対象ファイル: tools/live-check-runner（config/auth.json）
対象関数:    npm run test:jyu:smoke 等
対象URL:     dev/@13 prod
対象SS:      —
本番書込リスク: なし（read-only smoke）
GAS deploy 必要性: なし
Cloud Run deploy: なし
live-check 確認可否: 本作業自体が live-check 整備
人の確認作業: Chrome 起動 + Google ログイン（save-auth）
制度リスク:   なし
優先度:      中（全フェーズの検証前提＝enabler）
先にやる理由: これが無いと他フェーズを Playwright で検証できない
後回し可:    単体では製品価値を生まない。実装フェーズの「必要lleve-check」に内包できる
```

### C. KPI handler 正式 commit（dual-deploy parity 回復）

```
目的:        Editor 上にのみ存在する ?action=insuranceKpiSummary handler を Ver3_core.js に正式 commit
対象ファイル: Ver3_core.js（doGet に action 分岐を追加）
対象関数:    doGet / insuranceKpiSummary handler
対象URL:     KPI @15
対象SS:      JREC本体（読み取り）
本番書込リスク: 低（読み取り集計）。ただし実効化には KPI deployment への deploy が必須
GAS deploy 必要性: あり（KPI deployment＝最も危険な操作。@15 pin の解除を伴う）
Cloud Run deploy: 不要
live-check 確認可否: 可（KPI endpoint は ANONYMOUS read。HTTP GET で検証可）
人の確認作業: JBIZ Portal-13 連携の regression 確認
制度リスク:   低（KPI は経営指標。請求正誤に非該当）
優先度:      中（latent footgun の解消だが、実行自体が危険操作）
先にやる理由: 将来 KPI deploy で handler が消える事故を恒久解消
後回し可:    @15 pin を維持していれば当面顕在化しない。実行時に dual-deploy 慎重対応が必要なため、監査基盤を先に持つ方が安全
```

### D. 医師同意管理の自動チェック ★返戻リスク最大の論点

```
目的:        骨折・脱臼で医師同意が無い申請の返戻を防ぐ（SPEC §539/§557）
対象ファイル: Ver3_amounts.js / Ver3_core.js（区分判定・needCheckReason）/ 申請書摘要欄（Ver3_transferData.js）
対象関数:    calcOnePartAmount_V3_ / saveVisit_V3 / saveVisitFromWeb_V3 / V3TR_build*（摘要）
対象URL:     staff /exec
対象SS:      JREC本体（来院ヘッダ書込）+ 設定（同意データの置き場が未定）
本番書込リスク: 中（同意データの保存先を新設すると新規 write path が発生）
GAS deploy 必要性: あり
Cloud Run deploy: 帳票摘要に反映する場合は要検討
live-check 確認可否: 可（骨折ケース登録 → needCheckReason に「医師同意未確認」を assert）
人の確認作業: ★同意データの保存先（来院ヘッダ新列 / 専用シート / 摘要直入力）のオーナー決定が必須
制度リスク:   高（返戻・査定の直撃領域）
優先度:      高（ただし前提決定と監査基盤が要る）
先にやる理由: 返戻リスク #1 の直接低減
後回し可:    同意データモデルが未定（オーナー判断待ち）。監査ログ無しで制度ロジックを変えると事故時に追跡不能
```

### E. 監査ログ強化（汎用操作ログ基盤の新設）★推奨

```
目的:        来院登録/修正/削除・申請書生成・施術録生成を append-only の操作ログに残す
対象ファイル: Ver3_core.js（saveVisit_V3 / saveVisitFromWeb_V3 / 削除経路）/ 新規ログ writer
対象関数:    （新設）writeOperationLog_V3_ + 既存 V3TR_writeGenerationLog_ との統一
対象URL:     staff /exec（登録経路）
対象SS:      JREC本体（新シート _操作ログ を append-only）
本番書込リスク: ★最小（既存シートを一切変更せず、新シートに追記のみ。auto-create + frozen header）
GAS deploy 必要性: あり（登録経路にフック → staff deployment へ deploy）
Cloud Run deploy: 不要
live-check 確認可否: 可（テスト visitKey 登録 → _操作ログ に行追加を gviz/read-only で assert。要 test-data 削除運用）
人の確認作業: ログ列設計の最終確認・個人情報ログ禁止フィールドの遵守確認
制度リスク:   低（追加情報の記録のみ。算定ロジック不変）
優先度:      高（監査耐性・外販耐性を即時付与。D/F の安全な土台）
先にやる理由: 追記のみで事故リスク最小。以後の制度変更（D）を監査可能にする前提
後回し可:    返戻リスクの直接低減ではない（ただし D を安全にやるための前提）
```

### F. 申請書生成/施術録生成の安全確認フロー整備

```
目的:        不可逆出力（Drive保存・ログ書込）前の dry-run / preflight / 確認を強化
対象ファイル: Ver3_transferData.js（B案）/ Ver3_shuRecorder.js
対象関数:    V3TR_menuGenerateApplication_B / generateClaimApplicationBFromWeb_V3 / srGenerateDocument
対象URL:     スプレッドシートメニュー / staff /exec
対象SS:      JREC本体（_申請書生成ログ）+ Drive
本番書込リスク: 中（生成系は Drive 出力・ログ書込を伴う）
GAS deploy 必要性: あり
Cloud Run deploy: 不要（B案ロジックは変えない）
live-check 確認可否: 一部（生成は本番書込のため dry-run 経路の整備が前提）
人の確認作業: dry-run 結果の目視
制度リスク:   中（生成物が制度帳票）
優先度:      中（B案には既に Preflight 検証あり。完全な空白ではない）
先にやる理由: 事故ゼロの本丸（最も危険な不可逆操作）
後回し可:    既存 Preflight である程度カバー済み。監査ログが先にある方が dry-run 差分を追える
```

### G. 月次請求・施術録出力の院内運用フロー確定

```
目的:        毎月の請求・施術録出力の人手運用手順を文書化
対象ファイル: docs のみ（コード非該当）
対象関数:    —
本番書込リスク: なし
GAS deploy 必要性: なし
Cloud Run deploy: なし
live-check 確認可否: 非該当
人の確認作業: 院内運用の合意
制度リスク:   低
優先度:      低〜中（運用ドキュメント）
先にやる理由: 現場定着には有用
後回し可:    コード基盤が固まってからで良い
```

---

## 2. 既存実装 × 制度残課題の照合と実装順序

実装済み（前回棚卸し）: 30日ルール / 初検・再検・後療 / 同日複数case / 多部位逓減 / 長期減額75-50 / 月内上限 / 近接部位（保存ブロック）/ 申請書明細生成 / 施術録生成 / 抑制理由ログ。

残課題:
- **医師同意管理**: SPEC §539/§557 に規定あり、js 自動チェックなし。同意データの保存先が未定義。
- **監査ログ**: `_申請書生成ログ`（生成イベントのみ）。来院登録/修正/削除の汎用操作ログは未整備。

推奨実装順序（事故ゼロ・監査耐性を優先）:

```
Phase 1（次）  E. 監査ログ基盤新設（_操作ログ）        ← append-only・最小リスク・D/Fの土台
Phase 2        D. 医師同意チェック（非ブロッキング警告）  ← 返戻リスク直撃。Phase1で監査可能になった上で実施
                  ※ 着手前にオーナー決定: 同意データの保存先
Phase 3        F. 生成系 dry-run/preflight 強化          ← 監査ログで差分を追える状態で実施
Phase 4        C. KPI handler 正式 commit + 慎重 dual-deploy ← latent footgun の恒久解消
Phase 5        A/G. UI 導線整理・院内運用フロー確定        ← 制度基盤が固まってから
補助           B. live-check auth 更新                    ← Phase1 の検証時にまとめて実施（enabler）
```

理由: 返戻リスク最大は D だが、D は (1) 同意データモデルのオーナー決定待ち、(2) 制度ロジック変更で本番書込を伴うため、**監査ログ（E）が無い状態で D を入れると事故時に追跡不能**。E は追記のみで事故リスク最小、監査耐性・外販耐性を即時付与し、D を安全に載せる土台になる。よって E を先頭に置く。

---

## 3. 次フェーズ 推奨結論（1つに絞る）

```
推奨フェーズ:  E. 監査ログ基盤新設 — _操作ログ シート（append-only 汎用操作ログ）
理由:
  - 本番書込事故リスクが全候補で最小（既存シート不変・新シートに追記のみ）
  - 監査耐性・外販耐性（外販モデルに耐える）を即時付与＝ユーザー重視軸を直接満たす
  - 既存 V3TR_writeGenerationLog_（_申請書生成ログ）の実証済みパターンを踏襲＝設計リスク低
  - 後続の D（医師同意＝返戻リスク直撃）/ F（生成安全化）を「監査可能な基盤の上で」安全に実施できる
  - live-check で検証しやすい（テスト登録 → ログ行追加を read-only assert）

今回やる範囲（Phase 1 実装フェーズで行うこと）:
  - 新シート _操作ログ の新設（auto-create + frozen header）
  - writeOperationLog_V3_(ss, entry) ヘルパー新設（append-only）
  - saveVisit_V3 / saveVisitFromWeb_V3 / 来院削除経路 から呼び出しフック
  - 列案: 実行日時 / 操作種別(visit_create|visit_update|visit_delete|claim_generate|shuroku_generate) /
          実行経路(SheetsUI|Web|menu) / 患者ID / visitKey / 対象月 / kubun / needCheck /
          抑制理由要約 / 実行者(email 取得可否に依存) / 結果(OK|ERROR) / 詳細
  - 既存 _申請書生成ログ とは併存（生成系は両方に記録 or 段階統一）

今回やらない範囲:
  - 医師同意チェック（Phase 2）
  - 生成系 dry-run 強化（Phase 3）
  - KPI handler commit/deploy（Phase 4）
  - 既存シートの列変更・算定ロジック変更

対象repo:            C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1（main）
対象ファイル:        Ver3_core.js（フック）/ 新規 writer（Ver3_core.js または新ファイル）
対象関数:            saveVisit_V3 / saveVisitFromWeb_V3 / 削除経路 / writeOperationLog_V3_（新設）
対象URL:             staff /exec @17（登録経路）
対象スプレッドシートURL: https://docs.google.com/spreadsheets/d/1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM/edit
必要な確認:
  - 個人情報ログ禁止フィールド（氏名/住所/電話/生年月日/保険者番号/記号番号/被保険者情報）を _操作ログ に出さない
  - Session.getActiveUser().getEmail() が webapp(USER_ACCESSING) で取得可能かの実機確認
必要なlive-check:    auth 更新 → テスト visitKey 登録 → _操作ログ 行追加を gviz read-only で assert（test-data 削除運用込み）
必要な記録:          docs（本フェーズ実装記録）/ PROJECT_STATUS / SPEC（ログ仕様を §追記）/ Dashboard 反映
完了条件:
  - _操作ログ 新設 + 3 経路フック + live-check assert PASS
  - 既存シート無変更・回帰 smoke PASS
  - 個人情報ログ禁止フィールド非出力を確認
  - docs/SPEC/PROJECT_STATUS 記録 + commit/push + clean + 0-0
  - clasp push 前ゲート（git ls-files -d 空）確認
```

---

## 4. 次回 Claude 用 実装プロンプト案（Phase 1: 監査ログ基盤）

> 下記は次回の実装フェーズで使うプロンプト案。**本フェーズ（設計）では実装しない。**

```text
JREC-01 / jyu-gas-ver3.1 Phase 1: 監査ログ基盤（_操作ログ）実装。

# 前提・開始時確認（必須）
cd C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1
git status --short / git branch --show-current / git pull --ff-only
git log -1 --oneline / git rev-list --left-right --count HEAD...@{u} / git ls-files -d
- repo: main / 正本: GitHub dabu-pi/jyu-gas-ver3.1 main
- 本体SS: https://docs.google.com/spreadsheets/d/1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM/edit
- 設計根拠: docs/JREC-01_NEXT_PHASE_PLAN_2026-06-03.md §3

# 対象
- ファイル: Ver3_core.js（saveVisit_V3 / saveVisitFromWeb_V3 / 来院削除経路 / 新設 writeOperationLog_V3_）
- 参照パターン: Ver3_transferData.js V3TR_writeGenerationLog_（_申請書生成ログ の append-only writer）
- 新シート: _操作ログ（auto-create + frozen header / 既存シートは一切変更しない）

# 実装内容
1. writeOperationLog_V3_(ss, entry) を新設（append-only / auto-create / frozen header）
2. 列: 実行日時 / 操作種別 / 実行経路 / 患者ID / visitKey / 対象月 / kubun / needCheck / 抑制理由要約 / 実行者 / 結果 / 詳細
3. saveVisit_V3 / saveVisitFromWeb_V3 / 削除経路にフック（成功・失敗とも記録）
4. 個人情報ログ禁止フィールド（氏名/住所/電話/生年月日/保険者番号/記号番号/被保険者情報）は出力しない
5. 既存 _申請書生成ログ は併存（変更しない）

# 禁止事項
- 既存シートの列変更 / 算定ロジック変更 / application_template.xlsx 変更
- Cloud Run deploy / Secret Manager 変更
- 設計外の本番データ書込（テスト visitKey 以外の登録）

# 本番書込可否
- _操作ログ への append は許可（新シート追記のみ）
- 検証はテスト visitKey（例 hirayamaka_2999-12-31）で行い、終了後に test-data を削除

# dry-run / read-only 確認
- 実装後 clasp push 前ゲート: git update-index -q --refresh && git ls-files -d が空であること
- live-check-runner: auth 更新（Chrome 9222 + save-auth、single-writer 確認）→ npm run test:jyu:smoke（回帰）
- テスト登録 → gviz read-only で _操作ログ 行追加を assert → test-data 削除

# 実機確認URL
- staff /exec: https://script.google.com/macros/s/AKfycbxODNWJ.../exec
- _操作ログ gviz: https://docs.google.com/spreadsheets/d/1rXWkfAc.../gviz/tq?sheet=_操作ログ

# 記録先
- docs/JREC-01_PHASE1_AUDIT_LOG_<date>.md / PROJECT_STATUS.md / SPEC.md（ログ仕様 §追記）

# Dashboard 反映
- 人が workspace ルートで: de -ProjectId JREC-01 "JREC-01 Phase1 監査ログ <date>"

# commit / push
- clasp push（disk→GAS）→ commit（コード+docs）→ push。clean / 0-0 / missing 0 を確認

# 最終報告
- STATUS / REPO / 変更ファイル / live-check 結果 / _操作ログ 検証 / 個人情報非出力確認 /
  test-data 削除確認 / commit-push / 次フェーズ(D 医師同意) 申し送り
```

---

## 5. Dashboard 反映状況

```
実施: 未実施（自動書込せず）
理由:
  - de は PowerShell プロファイル関数で Claude 非対話シェルから呼べない（Get-Command de → NOT available）
  - sync スクリプト（append-runlog/sync-project）は想定パスに未発見
  - Run_Log/Projects/Task_Queue は single-writer 対象。並行 claude 稼働中につき回避
  - JREC-01 のローカル runlog/Projects 痕跡なし（Projects 未登録の可能性）
env: AIOS_DASHBOARD_SPREADSHEET_ID / AIOS_SERVICE_ACCOUNT_PATH は SET 済み（service account ファイルも存在）
人が実行するコマンド（workspace ルート）:
  de -ProjectId JREC-01 "JREC-01 次フェーズ設計確定 2026-06-03（docs only）"
判断基準: Projects に JREC-01 行があり 次アクション/最終更新日/補足 が更新されること。
         未登録なら de は行追加しない（[WARN] Skip: no auto-append）→ 先に行登録要否を人が判断。
```

---

## 6. 次回再開情報

```
repo:   C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1
branch: main / HEAD はこの commit（次回 git pull --ff-only）
本体SS: 【毎日記録】来店管理施術録ver3.1 / 1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM
次フェーズ: Phase 1 = E 監査ログ基盤（_操作ログ）。プロンプト案は本ファイル §4。
申し送り: Phase 2 = D 医師同意（着手前に「同意データ保存先」をオーナー決定）。
```
