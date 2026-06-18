# JBIZ事業トップ連携メモ（2026-06-18）

> JBIZ 名称・階層・役割分担の正本は hirayama-jyusei-strategy/docs/JBIZ_NAMING_AND_LAYER_RULES_2026-06-18.md。
> このプロジェクトはJBIZの一部。JBIZ起動ハブは軽量入口で、実務導線の本体は各事業トップに置く。
> このプロジェクトClaudeは JBIZ本体(gas/portal-gateway-v1.gs)を直接編集せず、ここに反映待ち情報を記録する。

## このプロジェクトのJBIZ上の位置づけ
- 事業トップ：接骨院トップ（保険請求管理）
- 表示区分：B
- 技術種別：GAS（clasp / GitHub main 正本）。柔整毎日記録システム（来院登録・区分判定・算定・療養費支給申請書生成）。GitHub `dabu-pi/jyu-gas-ver3.1` main 正本、clasp deploy(@N) は派生物。
- JBIZ起動ハブ表面に出すべきか：出さない（起動ハブには「接骨院トップ」への入口 1 つだけを置く）
- 理由：保険請求の実務導線（来院登録・月次申請・申請書生成）は本番書込・不可逆処理を含み数も多い。起動ハブは軽量リンク集に徹し、実務導線はすべて「接骨院トップ（=staff ホーム ?page=home）」に集約する。動的 KPI（保険来院数）は JBIZ Home 側で別 deployment（KPI @15）が担当しており、起動ハブで重い fetch はしない。

## この事業トップに置くべき機能
- 毎日使う入口：患者検索（?page=search）、来院記録登録（ホーム「来院記録」カード → 患者検索経由 → ?page=visitNew）
- たまに使う入口：月次申請一覧（?page=monthlyClaims）、月次申請詳細・申請書生成（?page=monthlyClaimDetail）、自費明細入力（?page=selfpay）、患者詳細（?page=detail）
- 今日見るもの：当月の保険来院サマリ、月次申請の「要確認（needCheck=true）」件数（※現時点では未集約。ホームに安全バナー「実務確認中」を表示中）
- KPI：当月保険来院数（insuranceKpiSummary / KPI deployment @15 が JBIZ Home へ供給）。当月対象月の保険来院件数（2026-06 時点は 0 件 = 正常）
- 人の確認待ち：(1) 月次申請の空状態 UI 改善（保険来院 0 件時の案内文・「詳細画面」説明・本番処理注意の一覧側表示・ボタン文言整理）(2) 申請書 Excel の印刷プレビュー目視（1ページ収まり・転帰丸囲み・罫線）
- 次にやること：月次申請画面の空状態・詳細説明・本番処理注意・ボタン文言整理（UIのみ・本番書込なし）

## JBIZ起動ハブ表面に出す直リンク候補（原則3〜5個）
- 接骨院トップ（保険請求管理）ホーム … staff `?page=home`（実務導線はここに集約。起動ハブ表面はこの 1 リンクのみ推奨）
- （任意・2個目を置く場合）患者検索 … staff `?page=search`
- （任意・3個目を置く場合）月次申請一覧 … staff `?page=monthlyClaims`
- ※ いずれも staff deployment 固定 URL（`getExternalPortalUrl_` / `Business_Links` / `?view=` resolver 経由で解決すること。deploymentId をハードコードしない）

## 事業トップ内に置く導線
- ホーム（?page=home）: 安全バナー「JREC-01 保険施術録 / 実務確認中 / 本体SS名」、来院記録カード（→患者検索経由）、月次申請カード、平山ビジネスポータルへ戻るリンク
- 患者検索（?page=search）→ 患者詳細（?page=detail）→ 来院記録追加（?page=visitNew）/ 自費明細入力（?page=selfpay）
- 月次申請一覧（?page=monthlyClaims）→ 月次申請詳細（?page=monthlyClaimDetail）→ Step1 転記データ生成 / Step3 申請書Excel生成（B案 Cloud Run・正ルート）
- 全 staff 7 page 共通: グローバルナビタブ（.web-nav）+「平山ビジネスポータルへ戻る」リンク（target=_blank）

## Homeから移した方がよい情報
- 保険来院数 KPI の「実数値表示」: JBIZ Home でサマリを見せ、ドリルダウン（患者別・月次の対象者一覧）は接骨院トップ ?page=monthlyClaims 側に置く
- 月次申請の「要確認件数」: Home では件数バッジのみ、実際の確認・処理は接骨院トップへ誘導
- 申請書生成・来院登録などの本番書込導線は Home に出さず、必ず接骨院トップ内に閉じる

## 開発者向け詳細に隠すもの
- GitHub: `github.com/dabu-pi/jyu-gas-ver3.1`（main が正本 / clasp deploy @N は派生物）
- local path: `C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1`
- branch: `main`（※ この repo の正本は main。workspace 共通の feature/auto-dev-phase3-loop ではない）
- Run_Log: workspace AIOS Dashboard 側（本体SS内には無し）。`de -ProjectId JREC-01 "..."` で反映
- PROJECT_STATUS: `gas-projects/jyu-gas-ver3.1/PROJECT_STATUS.md`（最終更新 2026-06-03 / UI実務化フェーズ1 実機確認 + ハンドオフ）
- Claude再開情報: `docs/JREC-01_HANDOFF_2026-06-03.md` §4 / `docs/JREC-01_現状確認_2026-06-03.md` §14。次回最初の作業 = 月次申請画面の空状態 UI 改善（UIのみ）。本体SS = 【毎日記録】来店管理施術録ver3.1（ID `1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM`）。scriptId `1LROlc63TPr4Y2uV3nT6tAwOc-T0bOFflF0aXNzTlLEM_C3QbydHxCTzH`

## JBIZ Claudeへ渡す反映情報
- 起動ハブへ追加したいボタン：「接骨院トップ（保険請求管理）」1 ボタン（staff `?page=home` へ）。重い fetch なし。
- 事業トップへ追加したい導線：起動ハブ → 接骨院トップ → 各実務導線（患者検索 / 月次申請 / 自費明細）。実務導線本体は staff ホーム内に集約済み。
- 注意すべきURL drift：JYU-GAS には用途の異なる 2 deployment があり、起動ハブ/Home リンクが古い deploymentId を向く drift に注意。
  - staff entry deployment（公開導線 / 全 page routing）= `unknown（要確認）`（現行 @19。staff の deploymentId は履歴に `AKfycbxODNWJ...` 系として記録されているが、本メモではハードコードせず resolver で解決すること）
  - KPI deployment（`?action=insuranceKpiSummary` のみ）= `unknown（要確認）`（@15 に pin 維持。HEAD code を deploy すると handler 消失 → JBIZ KPI 連携破壊）
- 外部URL / 内部view / resolver利用の注意：JBIZ Portal は deploymentId を直書きせず `getExternalPortalUrl_` / `Business_Links` の primary_url / `?view=` resolver を使って JYU-GAS staff URL を解決する。ScriptProperty / Business_Links 既存行は seed 関数変更だけでは更新されないため deploy 後に diagnose/update runtime 実行が必要。
- 重い処理：保険来院数 KPI 集計（insuranceKpiSummary）は KPI deployment（@15）が担当。起動ハブ表面では実行しない（Home で fetch）。申請書 Excel 生成は Cloud Run（`jrec-appgen-server`）への外部リクエストで重い・本番書込のため起動ハブ/Home に出さない。
- 人の確認事項：(1) 月次申請の空状態 UI 改善方針の確認 (2) 申請書 Excel 印刷プレビュー目視 (3) KPI handler を Ver3_core.js へ正式 commit するか（dual-deploy parity 回復の follow-up タスク）

## 次にこのプロジェクトClaudeがやること
- 月次申請画面（web-monthly-claims.html）の空状態・詳細説明・本番処理注意バナー・ボタン文言整理を UI のみで改善（本番書込・申請書生成・Drive出力・KPI deployment 変更はしない）
- staff deployment への clasp deploy 時は deploymentId を維持（`clasp deploy --deploymentId <既存ID>`）し、KPI deployment（@15 pin）は触らない。clasp push 前に `git ls-files -d` が空であることを確認
- follow-up: Editor 上にのみ存在する `?action=insuranceKpiSummary` handler を Ver3_core.js へ正式 commit（dual-deploy parity 回復）
