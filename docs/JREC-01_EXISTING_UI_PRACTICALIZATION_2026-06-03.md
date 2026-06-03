# JREC-01 既存 staff WebUI 実務化フェーズ1 — 2026-06-03

既存の JREC-01 staff WebアプリUI を「実務で確認できる状態」に近づけるためのUI棚卸し・最小限UI修正の記録。
**制度ロジック・保存処理・申請書生成処理・施術録生成処理は一切変更していない（UI/表示文言のみ）。**

## 0. 対象・正本

| 項目 | 値 |
|---|---|
| 実機確認URL（staff /exec @17） | https://script.google.com/macros/s/AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_Ilt8SS5P5zodfF2dnmKeqso8BL8hcinVEBrQ/exec |
| 本体SS | 【毎日記録】来店管理施術録ver3.1 / `1rXWkfAc_ppOfMV5Dxmb3maX9ORVrZbpSOX2Lz7RouZM` |
| コード正本 | GitHub dabu-pi/jyu-gas-ver3.1 main |
| deployment | staff @17（`AKfycbxODNWJ...`）/ KPI @15 pin（`AKfycbxNMVV...`／触らない） |
| doGet ルート | home / search / detail / visitNew / selfpay / monthlyClaims / monthlyClaimDetail / findMonths / b2Results / fixtureResults（既定=patientSearch） |

## 1. 確認方法と制約

- 既存UIはコードレビュー（HTML + Ver3_core.js doGet）で構造確認。
- **Playwright live-check は未実行**: auth.json が 2026-05-28 付（6日前・Google認証失効疑い）/ Chrome CDP 9222 未起動 / 別 claude プロセス稼働（single-writer 競合）。
- staff /exec は `access: MYSELF` のため認証なし read-only HTTP では中身を取得できない（ログインリダイレクト）。
- → 実画面の目視確認は**人手（Google ログイン）が必要**。本書 §6 に人手確認URL・手順を記載。

## 2. 既存UI棚卸し（10画面）

### ホーム（?page=home / web-home.html）
```
現状:        ナビ4タブ + カード（患者検索/来院記録=稼働, 施術録/設定/監査ログ=準備中, 月次申請=稼働）+ ポータル戻り
実務利用可否: ◯（入口として機能）
不足:        アプリ全体の「実務確認中」明示・本体SS明示・保存系注意が無かった
危険な導線:   なし（生成系への直リンクなし）
改善案:      「実務確認中」安全バナー追加（本体SS名・保存/生成は確認後に実行）
今回直すか:   ★直した（安全バナー追加）
後回し:      施術録カードの web 導線（現状 Sheets メニューのみ）は新規routeになるため後回し
理由:        入口に状態と注意を出すのが実務化の最優先・最小リスク
```

### 患者検索（?page=search / patientSearch.html）
```
現状:        氏名/ID 検索 → 患者詳細・自費明細導線。ナビ・ポータル戻りあり（AUDIT-003）
実務利用可否: ◯
不足:        特になし（読み取り中心）
危険な導線:   なし
改善案:      （任意）検索結果の視認性。今回は不要
今回直すか:   直さない
理由:        読み取り画面で本番書込リスクなし。現状で実務確認可能
```

### 患者詳細（?page=detail / web-patient-detail.html）
```
現状:        患者基本情報 + 来院履歴（読み取り）+ 来院追加/自費明細導線
実務利用可否: ◯
危険な導線:   なし（表示のみ）
今回直すか:   直さない
理由:        読み取り中心。実務確認に支障なし
```

### 新規来院登録（?page=visitNew / web-visit-new.html）
```
現状:        来院フォーム + 確認モーダル + 候補金額表示（needCheck=true 記録）
実務利用可否: △（実務可だが本番書込導線。保存は saveVisitFromWeb_V3）
不足:        保存前の最終確認（対象患者/施術日/区分/部位/金額）の明示強化余地
危険な導線:   保存ボタン（本番書込）。ただし既存で確認モーダルあり
今回直すか:   今回直さない（保存系の確認UIは保存ロジックに近く、誤って挙動を変えるリスク。フェーズ2で慎重に）
後回し:      ◯
理由:        保存ロジックに触れずに確認UIだけ強化するには慎重設計が要る。今回はホーム/生成導線の注意を優先
```

### 月次申請一覧（?page=monthlyClaims / web-monthly-claims.html）
```
現状:        年月入力 → 対象者一覧 → 詳細へ遷移（読み取り）
実務利用可否: ◯
危険な導線:   なし（一覧表示）
今回直すか:   直さない
理由:        読み取り画面。実務確認に支障なし
```

### 月次申請詳細（?page=monthlyClaimDetail / web-monthly-claim-detail.html）★最重要
```
現状:        日別明細 + 4生成ステップ（Step1 転記データ / Step2 NDJSON / Step3 B案Excel生成 / Step4 試験的PDF）
             各ボタンに confirm() + 個別 note。要確認フラグ時は警告表示
実務利用可否: △（実務可だが本番書込/Drive出力/ログ記録の集中地点）
不足:        生成エリア全体を覆う標準「本番処理注意」が無かった
危険な導線:   Step1〜4 すべて本番書込系。Step4(A案PDF)は停止ルートだが押下可能（試験的表示済）
改善案:      生成エリア先頭に標準安全文言（Drive出力/ログ記録・対象月/患者/保険者/金額を確認）追加
今回直すか:   ★直した（Step1直前に「本番処理注意」バナー追加）
後回し:      Step4(A案PDF)ボタンの非活性化・撤去はロジック近接のためフェーズ2で判断
理由:        生成導線の入口に統一注意を出すのが返戻・誤生成事故の最小化に直結
```

### 申請書生成導線
```
現状:        月次申請詳細の Step1〜3（B案 Cloud Run が正ルート）/ Sheets メニュー「帳票出力→申請書を出力」
実務利用可否: ◯（confirm + note + 今回の標準注意で多重化）
危険な導線:   Drive出力 + _申請書生成ログ 追記 + Cloud Run 実行
今回直すか:   ★標準注意を追加（上記詳細画面）
理由:        生成は不可逆出力を含む。注意の多重化が妥当
```

### 施術録生成導線
```
現状:        Web 導線なし（Sheets メニュー「施術録を出力」srGenerateDocument のみ）。ホームでは「準備中」
実務利用可否: ―（Web 未提供）
今回直すか:   直さない（Web route 新設は今回スコープ外）
後回し:      ◯（フェーズ2以降で web 導線検討時に安全注意込みで設計）
理由:        新規preview route を作らない方針。現状の「準備中」表示が正しい
```

### 自費明細（?page=selfpay / selfPayWeb.html）
```
現状:        自費明細入力。ナビ・ポータル戻りあり
実務利用可否: ◯（保険申請の本筋とは別系統）
今回直すか:   直さない
理由:        今回は保険対応導線の実務化が主目的。自費は別系統で現状維持
```

### ポータル戻り導線
```
現状:        web-home / patientSearch / selfPayWeb / monthlyClaims / monthlyClaimDetail / patient-detail / visit-new の
             7 staff page に「← 平山ビジネスポータルへ戻る」あり（AUDIT-003 @17 / target=_blank）
実務利用可否: ◯
今回直すか:   直さない（@17 で整備済み）
理由:        既に全 staff page に導線あり。重複追加不要
```

## 3. 今回のUI修正（最小・HTML/表示のみ）

| ファイル | 変更内容 | ロジック影響 |
|---|---|---|
| `web-home.html` | page-header 直下に「🩺 JREC-01 保険対応アプリ／実務確認中／本体SS名／保存・生成は確認後に実行」安全バナー追加 | なし（表示のみ） |
| `web-monthly-claim-detail.html` | Step 1 直前に「⚠️ 本番処理注意 — 申請書生成（Step1〜4）／Drive出力・ログ記録を伴う／対象月・患者・保険者・金額を確認」バナー追加 | なし（表示のみ） |

- 変更は 2 ファイル・計12行追加のみ。`git diff --name-only` に `.js` なし（HTML のみ）。
- 保存処理・算定ロジック・生成処理・テンプレートは一切不変。

## 4. 押してよい / 押してはいけないボタン（実務確認用）

**押してよい（読み取り・遷移のみ）:**
- ナビ各タブ / カード遷移 / 患者検索 / 患者詳細表示 / 月次申請一覧 / 月次申請詳細の表示 / ポータルへ戻る

**押す前に必ず内容確認（本番書込・不可逆）:**
- 月次申請詳細 Step1「申請書転記データを生成」（`buildMonthlyTransferData_V3` → 申請書_転記データ upsert）
- Step2「NDJSON を Drive に出力」（`exportClaimNdjson_V3` → Drive出力）
- Step3「申請書Excelを生成」（`generateClaimApplicationBFromWeb_V3` → Cloud Run + Drive + ログ）★正ルート
- Step4「申請書PDFを生成（試験的）」（`generateClaimApplication_V3` → A案・停止ルート。原則使わない）
- 新規来院登録の「保存」（`saveVisitFromWeb_V3` → 来院ケース/ヘッダ/施術明細 書込）

## 5. 本番書込リスク（今回の修正による増減）
- 増加: なし（注意表示の追加のみ。新たな書込導線は作っていない）
- 低減: 生成導線に標準注意を多重化 → 誤生成・確認漏れの抑止

## 6. deploy情報（2026-06-03 実施済み — staff @19）

```
実施:    ★実施済み（2026-06-03）。staff deployment のみ更新。KPI @15 は不変
deploymentId: AKfycbxODNWJ...（staff・不変）/ version @17 → @19 / URL 不変
KPI:     AKfycbxNMVV... @15 のまま（live read-only HTTP200 ok:true 確認）
検証:    clasp pull で editor HEAD を temp に非破壊 snapshot → 全 .js が editor=repo 一致・
         insuranceKpiSummary handler は editor に既に無し（version15 スナップショットのみ）→ フットガン非該当
rollback: clasp deploy --deploymentId AKfycbxODNWJ... --versionNumber 17
影響:    UI変更が staff /exec に反映済み（人手 Google ログインで目視）
```

**人が deploy する手順（並行 claude を止め、Google ログイン可能な状態で）:**
```powershell
cd C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1
git pull --ff-only
git update-index -q --refresh; git ls-files -d   # ← 空であること（空でなければ clasp push 禁止）
# clasp 認証済み前提。KPI deployment には絶対 deploy しない
clasp push --force
clasp deployments                                  # staff deploymentId(AKfycbxODNWJ...) を確認
clasp deploy --deploymentId AKfycbxODNWJNcCJVQnDXHzzWck237hnUIIXR_Ilt8SS5P5zodfF2dnmKeqso8BL8hcinVEBrQ -d "UI実務化フェーズ1: 安全バナー"
clasp deployments                                  # version が上がったこと(@17→@18)・URL不変を確認
```
rollback: `clasp deploy --deploymentId <staffId> --versionNumber 17`（@17 に戻す）。
KPI @15 は一切触らない（`clasp deploy --deploymentId AKfycbxNMVV... --versionNumber 15` を保持）。

## 7. live-check-runner 結果
```
実施:    Playwright 本体は未実施
理由:    auth.json 失効疑い(5/28) / CDP 9222 未起動 / 別 claude 並行(single-writer)
代替:    HTML コードレビューで構造・タグ整合を確認（挿入は単純 div・既存CSS非破壊）
人手確認: deploy 後に下記URLを Google ログインで開いて目視
```

## 7.5 ユーザー実機確認結果（2026-06-03・deploy @19 後）

- ホーム = ✅ OK（バナー表示 / 来院記録カード→患者検索 / ポータル戻り / 白画面なし）。
- 月次申請 = △ 一部のみ（当月 保険来院0件で対象者が出ず、詳細・生成エリア注意バナー・生成 confirm は未確認）。
- 次回: 月次申請の空状態・詳細説明・本番処理注意・ボタン文言の整理（UIのみ）。詳細は [`JREC-01_HANDOFF_2026-06-03.md`](./JREC-01_HANDOFF_2026-06-03.md)。

## 8. 人が確認するURL（deploy 後）
```
URL: https://script.google.com/macros/s/AKfycbxODNWJ.../exec?page=home
  確認: 「🩺 JREC-01 保険対応アプリ / 実務確認中 / 本体SS名」バナーが出る・白画面なし
URL: https://script.google.com/macros/s/AKfycbxODNWJ.../exec?page=monthlyClaims → 患者選択 → 詳細
  確認: 生成ステップ群の先頭に「⚠️ 本番処理注意」バナーが出る・各ボタンに confirm が残っている
  ※ Step1〜4 のボタンは確認用に「押さない」。表示のみ確認
```

## 9. 次に直すべきUI（フェーズ2候補）
1. 新規来院登録の「保存前確認」強化（対象患者/施術日/区分/部位/金額を確認パネル化。保存ロジックは不変で表示のみ）
2. 月次申請詳細 Step4（A案PDF・停止ルート）ボタンの非活性化または撤去（誤操作防止）
3. 施術録の web 導線（現状 Sheets メニューのみ）を安全注意込みで検討（新 route 設計が必要）

## 10. Dashboard 反映
```
実施:   未実施（de は非対話シェルで利用不可）
人が実行: cd C:\hirayama-ai-workspace\workspace
         de -ProjectId JREC-01 "JREC-01 既存UI実務化フェーズ1 2026-06-03"
判断基準: Projects に JREC-01 行があり 次アクション/最終更新日/補足 が更新されること
```
