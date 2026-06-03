# JREC-01 保険施術録UI — 自費アプリ(JREC-SF01)参考方針 — 2026-06-03

**方針:** 実務利用が進んでいる自費アプリ **JREC-SF01** の画面構成・導線・実務UX を参考に、保険施術録UI(JREC-01)を実務で使いやすく整える。
**参考にするのは UI / 導線 / 画面整理のみ。** 自費の金額計算・会計・予約・問診ロジックは保険側へ持ち込まない。
本フェーズはUI/表示/导线のみ変更。制度ロジック・保存処理・生成処理・テンプレートは不変。

## 0. 対象URL（実測）

| 区分 | URL / 値 |
|---|---|
| 保険 staff(@17) ★今回の正本確認URL | https://script.google.com/macros/s/AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_Ilt8SS5P5zodfF2dnmKeqso8BL8hcinVEBrQ/exec |
| 保険 本体SS | 【毎日記録】来店管理施術録ver3.1 / `1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM` |
| 自費 staff prod（参考・live-check config 実測） | https://script.google.com/macros/s/AKfycbyOtef10SuH7R1SaDVMBZS7L9yZIBYpEIVmNdS_fhz3hUtc1b0PSKvtzwRxQ6I43YObEA/exec |
| 自費 repo | `C:\hirayama-ai-workspace\workspace\gas-projects\jrec-sf01-selfpay`（main / HEAD `cda5f6b` @130/@131）|
| 自費 deployment（config 記録） | @67 staff UI / @57 public（repo HEAD は @130/@131）|

> 自費の URL/deployment は `tools/live-check-runner/projects/jrec-sf01/config.json` と repo PROJECT_STATUS から実測。推測なし。

## 1. 自費アプリUIの構造（参考にする要素）

自費アプリ(JREC-SF01)のUIアーキテクチャ（`home.html` / `index.html` / `styles.html` を実読）:

| 要素 | 内容 | 保険へ参考になる点 |
|---|---|---|
| **共通シェル** | `<?!= include('index') ?>`（ヘッダ+タブnav）+ `<?!= include('styles') ?>`（共通CSS）を全 staff ページに埋め込み | ★保険は include 未使用で各HTMLにnav重複。共通化すると保守性UP |
| **ヘッダ** | アプリ名「JREC-SF01 自費カルテ・会計」+ 院名サブ「平山接骨院」+ ポータル戻り帯 | アプリ名・院名の明示。保険も同様にできる |
| **タブnav** | ホーム/本日の受付・会計/患者一覧/＋新規患者登録/売上・レポート/📋問診票/📅予約管理/🌐公開予約/🔗URL発行。`CURRENT_PAGE`でactive強調 | active強調つきの一貫nav。保険にも有効 |
| **ホーム menu-card** | `menu-card` グリッド。日次主要動作（本日の受付・会計）を **primary card**(青強調)に | ★「今日まず押すボタン」を1つ強調する構成 |
| **ホームダッシュボード** | 予約状況サマリ（**PII無し・件数のみ**）+ 月間来院カレンダー（●印クリックで日次へ） | ★一目で当日状況。保険は「月次申請状況/要確認件数」に置換可 |
| **遷移方式** | `window.top.location.href`（iframe入れ子回避＝保険の target=_top と同義）| 既に保険も target=_top で対応済み |
| **PII規律** | 「集計値のみ表示。患者名・連絡先はトップに出さない」 | 保険の個人情報ログ禁止原則と一致。踏襲する |

## 2. 自費UI → 保険UI 差分表

```
自費アプリ画面:        ホーム（menu-card + 予約サマリ + 来院カレンダー）
保険で対応する画面:    web-home.html（?page=home）
流用できるUI:          menu-card primary 強調 / 共通シェル(include) / active付きnav / ホームダッシュボード枠
流用してはいけない処理: 予約サマリの予約ロジック・会計集計
保険側で追加すべき:    日次主要動作の primary card / 「月次申請 要確認件数」サマリ(将来) / アプリ名・院名・実務確認中明示
不足している導線:      来院記録カードが visitNew 直リンク=患者未指定で行き止まり → 患者選択経由が必要
今回直すか:            ★直す（来院記録カードを患者検索経由へ / 実務確認中バナー / 保険施術録表記）
後回し:                共通シェル(include)化・ホームダッシュボード実装は構造変更のため次フェーズ
理由:                  行き止まり導線と必須表示は最小修正で即解消。構造変更は別フェーズで安全に
```

```
自費アプリ画面:        本日の受付・会計（dailyCheckout）
保険で対応する画面:    （直接対応なし）保険は会計概念がない。日次は「来院登録」
流用できるUI:          「本日の一覧から個別へ」の日次起点の考え方
流用してはいけない処理: 会計・受付・売上ロジック（自費専用）
保険側で追加すべき:    （将来）当日来院一覧→来院登録/確認の日次ハブ
今回直すか:            直さない
後回し:                ◯
理由:                  保険に会計はない。日次ハブは構造設計が要るため次フェーズ
```

```
自費アプリ画面:        患者一覧・検索（list）
保険で対応する画面:    患者検索（?page=search / patientSearch.html）
流用できるUI:          検索→詳細→入力の一直線導線
流用してはいけない処理: なし（読み取り）
保険側で追加すべき:    検索結果から「来院登録」への直接導線（現状は詳細経由）
今回直すか:            今回は最小（ホーム来院記録カードを検索へ寄せることで補完）
後回し:                検索結果行への来院登録ボタンは次フェーズ
理由:                  既存検索は実務可。導線強化は段階的に
```

```
自費アプリ画面:        患者詳細（patient-detail）
保険で対応する画面:    患者詳細（?page=detail / web-patient-detail.html）
流用できるUI:          基本情報＋履歴＋アクションボタンの並び
流用してはいけない処理: 会計・未収表示
保険側で追加すべき:    来院記録/月次申請への自然な導線（既に来院追加ボタンあり）
今回直すか:            直さない
理由:                  既存で実務可
```

```
自費アプリ画面:        来院・カルテ入力（visit-form）
保険で対応する画面:    新規来院登録（?page=visitNew / web-visit-new.html）
流用できるUI:          patientId 必須を前提に患者選択から入る導線設計
流用してはいけない処理: 自費メニュー・会計・金額確定ロジック（保険は算定ロジックが別）
保険側で追加すべき:    保存前確認（対象患者/施術日/区分/部位/金額）の表示強化
今回直すか:            今回は導線のみ（ホームから患者選択経由に）。保存前確認UIは次フェーズ
理由:                  保存ロジック近接の確認UIは慎重設計が必要
```

```
自費アプリ画面:        売上・レポート（reports/monthlyReport/...）
保険で対応する画面:    月次申請一覧/詳細（monthlyClaims/monthlyClaimDetail）
流用できるUI:          月次集計の見せ方・KPIカード
流用してはいけない処理: 売上・未収・会計集計（自費専用）
保険側で追加すべき:    生成導線の本番処理注意（Drive出力/ログ記録の明示）
今回直すか:            ★直す（月次申請詳細の生成エリアに本番処理注意バナー）
理由:                  生成は不可逆出力。注意の多重化が返戻・誤生成事故を抑止
```

```
自費アプリ画面:        予約管理/公開予約/問診票/URL発行
保険で対応する画面:    （対応なし）
流用できるUI:          なし（保険ワークフロー外）
流用してはいけない処理: 予約・問診・公開URLロジック全般
保険側で追加すべき:    なし
今回直すか:            直さない
理由:                  自費専用機能。保険施術録に混ぜない
```

```
自費アプリ画面:        ポータル戻り（include index 内）
保険で対応する画面:    各 staff page の inline ポータル戻り帯（AUDIT-003 @17）
流用できるUI:          include で1箇所管理する方式
流用してはいけない処理: なし
保険側で追加すべき:    （将来）共通シェル化で1箇所管理
今回直すか:            直さない（@17で全page整備済）
理由:                  既に全 staff page に導線あり
```

### 保険側で必要な制度項目（自費には無い・UIで意識する）
visitKey / caseKey / 患者ID / 施術日 / 初検・再検・後療 / 受傷日 / 負傷原因 / 部位 / 医師同意 / 近接部位チェック / 多部位逓減 / 長期減額 / 月次申請 / 施術録 / 申請書 / 要確認理由 / 監査ログ。
→ これらは**制度ロジック側で実装済み/別フェーズ**。今回のUIフェーズでは「表示・導線・注意」に留め、ロジックには触れない。

## 3. 今回のUI修正（最小・HTML/导线/文言のみ）

| ファイル | 変更 | ロジック影響 |
|---|---|---|
| `web-home.html` | ① 「🩺 JREC-01 保険施術録／実務確認中／本体SS名／保存・生成は確認後に実行」安全バナー追加 ② 「来院記録」カードの行き止まり导线修正（`?page=visitNew` 直リンク→`?page=search` 患者選択経由、desc更新） | なし（表示・href のみ） |
| `web-monthly-claim-detail.html` | 申請書生成 Step1 直前に「⚠️ 本番処理注意／Drive出力・ログ記録を伴う／対象月・患者・保険者・金額を確認」バナー追加 | なし（表示のみ） |

- 変更 2 ファイル・計 ~14行。`git diff --name-only` に `.js` なし。
- 自費の**会計・予約・問診・金額ロジックは一切流用していない**（参考は画面構成・导线・注意表示・PII規律のみ）。

## 4. 押してよい / 押してはいけないボタン
**押してよい（読み取り・遷移）:** ナビ各タブ / カード遷移 / 患者検索 / 患者詳細 / 月次申請一覧 / 月次申請詳細表示 / ポータル戻り。
**押す前に必ず内容確認（本番書込・不可逆）:**
- 月次申請詳細 Step1 転記データ生成 / Step2 NDJSON出力 / Step3 申請書Excel生成(B案・正ルート) / Step4 PDF(試験的・A案停止ルート)
- 新規来院登録の保存（`saveVisitFromWeb_V3`）

## 5. 本番書込リスク
- 増加: なし（注意表示・导线修正のみ。新書込導線なし）。
- 低減: 行き止まり导线の解消 + 生成導線への本番処理注意多重化。

## 6. deploy情報（今回は見送り）

```
実施: 見送り（staff @17 未更新）
理由:
  - 別 claude プロセス(PID 7980 稼働)→ workspace CLAUDE.md single-writer
    「同一 Apps Script project への clasp push/deploy は直列のみ」回避
  - clasp push は Editor 上の未commit KPI handler を上書きするフットガン(@15 は v15 pin で live 無影響)
  - deploy 後検証が auth 失効・MYSELF access で自動不可（人手 Google ログイン要）
影響: UI変更は GitHub main に反映済み。staff /exec @17 には未反映（deploy 後に表示される）
旧URL維持: staff deploymentId を維持して deploy すれば URL 不変（@17→@18）
```

**人が deploy する手順（並行 claude を止め、clasp 認証済みで）:**
```powershell
cd C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1
git pull --ff-only
git update-index -q --refresh; git ls-files -d        # 空であること（空でなければ clasp push 禁止）
clasp push --force
clasp deploy --deploymentId AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_Ilt8SS5P5zodfF2dnmKeqso8BL8hcinVEBrQ -d "UI実務化: 保険施術録バナー+导线"
clasp deployments                                       # version 上昇(@17→@18)・URL不変を確認
# KPI deployment(AKfycbxNMVV...) には絶対 deploy しない（@15 v15 pin 維持）
```
rollback: `clasp deploy --deploymentId AKfycbxODNWJ... --versionNumber 17`

## 7. live-check-runner 結果
```
実施: Playwright 本体は未実施
理由: auth.json 失効疑い(5/28・6日前) / CDP 9222 未起動 / 別 claude 並行(single-writer)
代替: 保険・自費の HTML/doGet をコードレビュー。挿入は単純 div・既存CSS非破壊
人手: deploy 後に下記URLを Google ログインで目視
```

## 8. 人が確認するURL（deploy 後）
```
URL: https://script.google.com/macros/s/AKfycbxODNWJ.../exec?page=home
  確認: 「🩺 JREC-01 保険施術録 / 実務確認中 / 本体SS名」バナー / 「来院記録」カードが患者検索へ遷移 / 白画面なし
URL: https://script.google.com/macros/s/AKfycbxODNWJ.../exec?page=monthlyClaims → 患者選択 → 詳細
  確認: 生成エリア先頭に「⚠️ 本番処理注意」バナー / 各生成ボタンに confirm が残る
  ※ 生成ボタン(Step1〜4)は確認用に押さない。表示のみ
```

## 9. 次に直すべきUI（フェーズ2候補・自費構造の本格移植）
1. **共通シェル化**: 自費 `include('index')`/`include('styles')` 方式を保険へ導入し nav/ヘッダ/ポータル戻りを1箇所管理（doGet は createTemplateFromFile 維持・include 追加）。
2. **ホームダッシュボード**: 自費の「予約状況サマリ＋来院カレンダー」を、保険版「**月次申請 要確認件数 / 当月来院サマリ（PII無し）**」に置換実装。
3. **保存前確認UI強化**: 新規来院登録に「対象患者/施術日/区分/部位/金額」確認パネル（保存ロジック不変・表示のみ）。
4. Step4(A案PDF・停止ルート)ボタンの非活性化/撤去（誤操作防止）。

## 10. Dashboard 反映
```
実施: 未実施（de 非対話シェル不可）
人が実行: cd C:\hirayama-ai-workspace\workspace
         de -ProjectId JREC-01 "JREC-01 自費UI参考 保険施術録UI実務化 2026-06-03"
判断基準: Projects に JREC-01 行があり 次アクション/最終更新日/補足 が更新されること
```
