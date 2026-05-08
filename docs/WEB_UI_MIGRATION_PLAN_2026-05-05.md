# 柔整保険申請書 Ver3.1 Web UI 移行計画

作成日: 2026-05-05  
最終更新: 2026-05-05 (WEB-1B)  
フェーズ: WEB-1B 完了  
対象プロジェクト: JREC-01 (jyu-gas-ver3.1)

---

## 1. 現状のスプレッドシート構成

### 主要シート

| シート名 | 定数 (`SHEETS.*`) | 役割 |
|---|---|---|
| 設定 | `settings` | 単価・設定値一元管理 |
| 来院ケース | `cases` | 来院記録・算定ロジック実行結果 |
| 患者マスタ | `master` | 患者基本情報（氏名・フリガナ・生年月日等） |
| 患者画面 | `ui` | スプレッドシート操作UI（入力・確認パネル） |
| 施術明細 | `detail` | 部位・施術内容の明細 |
| 来院ヘッダ | `header` | 来院単位の集計・監査列 |
| 初検情報履歴 | `history` | 初検情報の履歴ログ |
| 保険者情報 | `insurer` | 保険者・被保険者情報 |
| 自費明細 | `selfPayDetail` | 自費メニュー明細（2026-03以降正本） |

### スプレッドシートの役割（維持方針）

- スプレッドシートは「データベース兼帳票基盤」として継続運用する
- Web UI はスプレッドシートを操作するフロントエンドとして追加する（置き換えではない）
- 既存の算定ロジック・帳票出力は GAS 関数のまま維持する

---

## 2. 現状の GAS 構成

### ファイル構成（clasp 管理）

| ファイル | 役割 |
|---|---|
| `Ver3_core.js` | 来院登録・区分判定・算定ロジック・Web App エントリポイント |
| `Ver3_amounts.js` | 金額計算（初検料・再検料・逓減・長期減額） |
| `Ver3_outputManager.js` | 月次出力フォルダ管理 |
| `Ver3_patientPicker.js` | 患者選択ダイアログ（スプレッドシート UI 用） |
| `Ver3_shuRecorder.js` | 施術録ドキュメント生成 |
| `Ver3_smokeTest.js` | スモークテスト |
| `Ver3_test.js` | 単体テスト |
| `Ver3_transferData.js` | 申請書データ転記 |

### Web App 設定（appsscript.json）

```json
{
  "webapp": {
    "executeAs": "USER_ACCESSING",
    "access": "MYSELF"
  }
}
```

> `access: MYSELF` のため、外部公開はしていない（本人アクセスのみ）。

### doGet ルーティング（Phase WEB-1 時点）

| `page=` | HTML | 状態 |
|---|---|---|
| `search`（デフォルト） | `patientSearch.html` | 稼働中 |
| `selfpay` | `selfPayWeb.html` | 稼働中 |
| `home` | `web-home.html` | WEB-1 追加 |
| `detail` | `web-patient-detail.html` | WEB-1 追加 |

---

## 3. Web UI 化の基本方針

### 原則

1. **既存スプレッドシート運用を壊さない**: Web UI は「並行入口」として追加する
2. **既存 GAS 関数を壊さない**: Web UI 用の新関数は末尾追加。既存関数の改修は最小限
3. **算定ロジックは後回し**: 金額計算・申請書生成ロジックの改修は WEB-3 以降
4. **読み取り優先**: WEB-1 は読み取り専用。書き込みは WEB-2 以降
5. **個人情報ログ禁止**: 氏名・住所・電話・生年月日・保険者番号は Logger に出力しない

### スプレッドシートを DB として残す理由

- 既存帳票テンプレート（Google ドキュメント）が Sheets 参照で動作している
- 算定ロジックのデバッグ・確認作業がスプレッドシート上で行われている
- 月次の保険請求業務は既存シート操作で完結している
- DB 化のコストとリスクに対してメリットが現時点では不十分

---

## 4. 将来的に DB 化する場合の考慮点

- `SHEETS.cases` の列定義が DB スキーマの代わり。移行前に正規化が必要
- `HEADER_COLS` 定数（来院ヘッダ列名）を DB テーブル設計の参考にする
- 患者マスタは個人情報保護の観点で暗号化・アクセス制御が必要
- 算定ロジック（逓減・長期減額・30日ルール）は DB 移行後も GAS 関数として継続するか、別サービスに移植するかを検討

---

## 5. フェーズ設計

### Phase WEB-1B（完了: 2026-05-05）— 入口整理・設計固め

- [x] `patientSearch.html` に「← Web ホームへ」リンク追加
- [x] `docs/PHASE_WEB2_VISIT_CREATE_DESIGN_2026-05-05.md`: `saveVisitFromWeb_V3` / `getPrevVisitData_V3` の設計
- [x] `web-home.html` は当面サブ入口（デフォルト URL は変更しない）
- [x] デフォルト route は `search` のまま維持（実地テスト済み導線を保護）
- [x] `web-home.html` default 化の条件を定義

**default route 方針:**

```
現在: page=search (patientSearch.html) がデフォルト → 維持
条件リスト(§web-home default化条件)が揃ったら page=home に変更する
```

**既存稼働導線（保護対象）:**

```
patientSearch.html → 患者選択 → setPatientAndDate_V3 → selfPayWeb.html → 自費明細保存
```

この導線は実地テスト済み。Phase WEB-2 以降も変更しない。

**web-home を default にする条件（将来）:**
1. `web-home.html` 実機確認 PASS
2. `patientSearch.html` にホームリンク追加済み（WEB-1B で実施）
3. `selfPayWeb.html` 既存導線が壊れていない
4. `web-patient-detail.html` 実機確認 PASS
5. 来院登録の基本導線が安定（Phase WEB-2 完了後）
6. 現場スマホ操作で実績

---

### Phase WEB-1（完了: 2026-05-05）

目標: 読み取り専用 Web UI の土台を作る

- [x] `doGet(e)` に `home` / `detail` ルート追加
- [x] `web-home.html`: ナビゲーションハブ
- [x] `web-patient-detail.html`: 患者詳細・来院履歴（読み取り専用）
- [x] `getPatientDetail_V3(patientId)`: 詳細データ取得 GAS 関数
- [x] `patientSearch.html`: 「患者詳細を見る」リンク追加

### Phase WEB-2（次フェーズ候補）

目標: Web から来院記録を新規登録できるようにする

対象:
- 患者選択 → 来院日選択 → 区分候補表示
- 部位・施術内容入力
- 保存前確認 → visitKey 発行
- 監査ログ記録

### Phase WEB-3（その次）

目標: Web から施術録・申請書生成へ進む

対象:
- 月次申請対象者一覧
- 申請書プレビュー
- 施術録裏面プレビュー
- 抑制理由表示
- 申請前チェック
- PDF / 印刷導線

---

## 6. 個人情報ログ禁止方針

### Logger.log に出力禁止のフィールド

- 氏名 (`name`, `患者氏名`)
- 住所 (`address`, `住所`)
- 電話番号 (`phone`, `電話番号`)
- 生年月日 (`birthday`, `生年月日`)
- 保険者番号
- 記号番号
- 被保険者情報

### ログに残してよいフィールド

- `patientId`（患者ID）
- `visitKey`（患者ID + 日付）
- `caseKey`（エピソードキー）
- `action`（操作種別）
- `result`（件数・成否）
- `reasonCode`（抑制理由コード）

---

## 7. 監査ログ方針（将来実装）

- 書き込み操作（来院登録・修正・申請書生成）は `初検情報履歴` に準じた別シートへ記録
- 記録内容: `visitKey`, `action`, `actorId`（GAS ユーザー), `timestamp`, `result`
- 個人情報は記録しない

---

## 8. 申請書出力ロジックを WEB-3 に後回しにする理由

- 既存の `srGenerateDocumentCombo` は Google ドキュメントのテンプレート操作を含む複雑な処理
- 帳票レイアウトは厚生労働省の様式に準拠しており、HTML/PDF では再現困難
- 既存帳票をそのまま使い続ける方が安全・確実
- Web UI では「申請書生成ボタン」を押すと既存 GAS 関数を呼び出す形にする

---

## 9. 既存帳票・既存シートを壊さない実装ルール

- `doGet(e)` の既存ルート（`search`, `selfpay`）は一切変更しない
- 新しい HTML ファイルは既存ファイルと命名が衝突しないよう `web-` プレフィックスを使う
- シートの列追加・削除・名称変更は行わない
- 新しい GAS 関数は `Ver3_core.js` 末尾の専用セクションに追記する
