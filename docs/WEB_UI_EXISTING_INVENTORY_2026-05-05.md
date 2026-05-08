# JREC-01 既存 Web UI 棚卸し

作成日: 2026-05-05  
目的: Phase WEB-2 着手前の現状整理  
調査範囲: `gas-projects/jyu-gas-ver3.1` のみ

---

## 1. HTML ファイル一覧

| ファイル | 種別 | 用途 | 呼び出し元 | 状態 |
|---|---|---|---|---|
| `patientSearch.html` | Web App | 患者検索・選択 | `doGet` page=search (デフォルト) | **稼働中・実地テスト済み** |
| `selfPayWeb.html` | Web App | 自費明細入力（Web版） | `doGet` page=selfpay | **稼働中・実地テスト済み** |
| `selfPayDialog.html` | SS ダイアログ | 自費明細入力（スプレッドシート版） | `openSelfPayDialog_V3()` | **稼働中（スプレッドシートUI）** |
| `web-home.html` | Web App | ナビゲーションハブ | `doGet` page=home | Phase WEB-1 追加（実機未確認） |
| `web-patient-detail.html` | Web App | 患者詳細・来院履歴（読み取り専用） | `doGet` page=detail | Phase WEB-1 追加（実機未確認） |

### 補足

- `selfPayDialog.html` は `SpreadsheetApp.getUi().showModalDialog()` 経由で開くためスプレッドシートからのみ利用可能。Web App の doGet とは完全に別系統。
- `selfPayWeb.html` と `selfPayDialog.html` は**同じ自費明細機能の別フロントエンド**。保存先は同一（自費明細シート）。

---

## 2. doGet ルート一覧（現在）

| page= | HTML | 既存/新規 | 用途 | 壊してはいけない理由 |
|---|---|---|---|---|
| `search` (デフォルト) | `patientSearch.html` | **既存** | 患者検索→選択→自費明細 | スマホ運用の入口。実地テスト済み |
| `selfpay` | `selfPayWeb.html` | **既存** | 自費明細入力 | `patientSearch.html` からの遷移先 |
| `home` | `web-home.html` | WEB-1 追加 | ナビゲーションハブ | — |
| `detail` | `web-patient-detail.html` | WEB-1 追加 | 患者詳細・来院履歴 | — |

### 注意

`web-home.html` は `?page=home` を明示しないと開かない。デフォルト URL はいまも `patientSearch.html`。
Phase WEB-2 前に「Web ホームをデフォルトにするか」を決める必要がある。

---

## 3. 既存 Web 導線（稼働中）

```
スマホ → Web App URL（?page=search がデフォルト）
↓
patientSearch.html
  → キーワード入力 → searchPatients_V3(keyword)
  → 患者カードをタップ
  → setPatientAndDate_V3(patientId)   ← 患者画面シート B2/B4 に書き込む
  → 「自費明細入力 →」リンク
↓
selfPayWeb.html（?page=selfpay&visitKey=xxx）
  → getSelfPayMenuMaster_V3()  メニューマスタ読み込み
  → getSelfPayDataByVisitKey_V3(visitKey)  既存明細読み込み
  → 項目入力 → 保存
  → saveSelfPayDetailsFromDialog_V3()  自費明細シートに保存
```

この導線は**実地テスト済み・スマホ運用中**。Phase WEB-2 以降も破壊しないこと。

---

## 4. Phase WEB-1 追加導線（新規・未実機確認）

```
?page=home → web-home.html
  → 「患者検索」カード → ?page=search （patientSearch.html へ）

patientSearch.html（選択後パネル）
  → 「患者詳細を見る →」ボタン（Phase WEB-1 追加）
  → ?page=detail&patientId=xxx → web-patient-detail.html
     → getPatientDetail_V3(patientId)
        → 患者マスタから基本情報
        → 来院ヘッダから直近10件の来院履歴
```

Phase WEB-1 の追加は既存導線を壊していないが、**接続が一方通行**：
- `patientSearch.html` → `web-patient-detail.html` は接続済み
- `web-patient-detail.html` → `selfPayWeb.html` への導線は含まれている（今日の visitKey で自費明細リンクを生成）
- `web-home.html` は現状、デフォルト URL からは到達できない（`?page=home` 指定が必要）

---

## 5. 既存関数の確認結果

### Web API として再利用可能（JSON 返却・副作用の説明付き）

| 関数 | 行番号 | 副作用 | Phase WEB-2 での再利用 |
|---|---|---|---|
| `searchPatients_V3(keyword)` | 5389 | なし（読み取りのみ） | ✓ そのまま使う |
| `setPatientAndDate_V3(patientId)` | 5499 | 患者画面シート B2/B4 に書き込む | △ シート書き込み依存のため要検討 |
| `getSelfPayDataByVisitKey_V3(visitKey)` | 5548 | なし（読み取りのみ） | ✓ そのまま使う |
| `saveSelfPayDetailsFromDialog_V3(vk, items, ctx)` | 5267 | 自費明細シートに保存 | ✓ 自費保存として使う |
| `getSelfPayMenuMaster_V3()` | 4522 | なし（JBIZ SS 読み取り） | ✓ 自費メニューマスタとして使う |
| `getCurrentVisitKey_V3()` | 5205 | なし（患者画面シート読み取り） | △ シート依存のため Web からは不安定 |
| `getPatientDetail_V3(patientId)` | 5808 | なし（読み取りのみ） | ✓ WEB-1 追加済み |
| `loadPrevSelfPayToDialog_V3()` | 5606 | PropertiesService 書き込み + ダイアログ表示 | × Web App からは呼べない |
| `getAndClearPrevSelfPayItems_V3()` | 5765 | PropertiesService 読み取り・削除 | △ Web 版では別実装が必要 |

### スプレッドシートUI 専用（Web App からそのまま呼べない）

| 関数 | 行番号 | 理由 |
|---|---|---|
| `saveVisit_V3()` | 1095 | 患者画面シートを全面的に読んで来院ケースに保存。Web 化には JSON 引数版の新関数が必要 |
| `autofillFromPreviousVisit_V3()` | 1864 | 患者画面シートへの書き込み前提。Web 化には前回データを JSON で返す版が必要 |
| `openSelfPayDialog_V3()` | 5162 | `SpreadsheetApp.getUi().showModalDialog()` 依存 |
| `refreshKeikaHistoryUI_V3()` | 1745 | 患者画面シートへの書き込み |
| `updateProgressFromUI_V3()` | 2539 | 患者画面シートから読んで来院ケースを更新 |

---

## 6. keikaNow / shoken / 経過履歴 の関係

### スプレッドシート UI 上の配置

| セル範囲 | 意味 | 種別 |
|---|---|---|
| `A16:B20` (case1_keikaNow) | 今回の経過（入力） | **入力用** |
| `A23:B28` (case1_shoken) | 所見（入力） | **入力用** |
| `D23:G28` (case1_keikaHistory) | 過去5件の経過履歴 | **表示専用** |
| `A40:B44` (case2_keikaNow) | ケース2 今回の経過 | **入力用** |
| `A47:B52` (case2_shoken) | ケース2 所見 | **入力用** |
| `D47:G52` (case2_keikaHistory) | ケース2 過去5件 | **表示専用** |

### 帳票での参照先

- `shoken` → 来院ケースシートの「所見」列に保存 → **施術録裏面に反映**
- `keikaNow` → 来院ケースシートの「経過_今回」列に保存 → **施術録裏面に反映**
- `keikaHistory` → 来院ケースシートから過去5件を取得して表示専用で表示（保存先ではない）

### Web UI 化するときの扱い

Phase WEB-2 で来院登録を実装するとき、`shoken` / `keikaNow` の入力フォームが必要。
ただし Phase WEB-2 スコープ（来院日・区分・部位・施術内容）より複雑なため、**別途 Phase WEB-2.5 として切り出す**のが安全。

---

## 7. Phase WEB-1 は延長か並列か

**判断: 部分延長 + 部分並列**

| 項目 | 判断 |
|---|---|
| `patientSearch.html` → `web-patient-detail.html` 接続 | **延長**（既存画面に新しいリンクを追加） |
| `web-home.html` 追加 | **並列**（デフォルト URL は変わっていない。`?page=home` 指定のみでアクセス可能） |
| 既存の `patientSearch.html` → `selfPayWeb.html` 導線 | **変更なし**（既存ルートを壊していない） |

`web-home.html` は接続されているが「入口」として使うには `?page=home` を知っている必要がある。
Phase WEB-2 前に**デフォルト URL の方針を決める**のが推奨。

---

## 8. Phase WEB-2 に進む前の注意点

### 必ず決めること

1. **デフォルト URL の方針**
   - いまのまま `patientSearch.html` をデフォルトにするか
   - `web-home.html` をデフォルトにして、そこから `patientSearch.html` に飛ぶかを決める

2. **来院登録（Web化）のアーキテクチャ**
   - `saveVisit_V3()` はシート依存が深すぎるため、そのまま Web から呼べない
   - `saveVisitFromWeb_V3(params)` のような JSON 引数版新関数が必要
   - 設計は Phase WEB-2 着手前に Markdown で固めること

3. **`setPatientAndDate_V3()` の位置づけ**
   - 現在この関数はシートの B2/B4 に書き込む副作用がある
   - Web UI から来院登録するとき、シート書き込みを経由するかどうかを決める
   - シートを経由しない設計にすれば既存スプレッドシート運用との干渉を防げる

### 実機確認が必要な箇所

- `web-home.html` が GAS Web App で正常表示されるか
- `web-patient-detail.html` が患者詳細・来院履歴を正しく表示するか
- `patientSearch.html` の「患者詳細を見る」ボタンが正しい URL で遷移するか

### 壊してはいけない既存導線

- `patientSearch.html` → `selfPayWeb.html` → 自費明細保存
- スプレッドシートメニュー → `openSelfPayDialog_V3()` → `selfPayDialog.html`
- `saveVisit_V3()` → 来院ケース + 来院ヘッダへの保存

---

## 9. 再利用候補（Phase WEB-2 以降）

### そのまま使える

- `searchPatients_V3(keyword)` — 患者検索（副作用なし）
- `getSelfPayMenuMaster_V3()` — 自費メニューマスタ取得
- `getSelfPayDataByVisitKey_V3(visitKey)` — 自費明細読み込み
- `saveSelfPayDetailsFromDialog_V3()` — 自費明細保存
- `getPatientDetail_V3(patientId)` — 患者詳細 + 来院履歴

### 新関数が必要

| 必要な機能 | 既存 | 新関数案 |
|---|---|---|
| 来院登録（Web から JSON で保存） | `saveVisit_V3()`（シート依存） | `saveVisitFromWeb_V3(params)` |
| 前回来院の引き継ぎデータ取得 | `autofillFromPreviousVisit_V3()`（シート書き込み） | `getPrevVisitData_V3(patientId, treatDate)` |
| keikaNow / shoken 取得 | スプレッドシートUIからのみ読み取り | `getCaseTextData_V3(visitKey)` |

---

## 10. 推奨方針

```
1. Phase WEB-2 着手前に実機確認を実施する
   （web-home.html / web-patient-detail.html が正常表示されるか）

2. デフォルト URL の方針を決める
   （web-home.html をデフォルトにするかどうか）

3. 来院登録の設計ドキュメントを Markdown で作成してから実装する
   （saveVisitFromWeb_V3 の引数設計・バリデーション方針を先に決める）

4. 保険側の来院登録（keikaNow / shoken 入力）は Phase WEB-2.5 として切り出す
   （Phase WEB-2 では部位・区分・施術内容を中心にする）

5. 既存スプレッドシート運用と並行できるよう、Web UI は「並行入口」として設計する
```
