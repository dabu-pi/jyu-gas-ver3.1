# Phase WEB-2 来院登録 設計ドキュメント

作成日: 2026-05-05  
フェーズ: WEB-1B（設計固め）  
実装対象: `saveVisitFromWeb_V3` / `getPrevVisitData_V3`

---

## 1. 既存関数の問題点

### `saveVisit_V3()` がそのまま Web から呼べない理由

`saveVisit_V3` は `SpreadsheetApp.getActive()` 前提で設計されており、**患者画面シート（UI シート）から全データを読み取る**構造になっている。

**シート依存している読み取り箇所:**

| セル | 内容 | 定数 |
|---|---|---|
| C2 | 患者ID | `UI.patientId` |
| B4 | 来院日 | `UI.treatDate` |
| B5 | ジム会員フラグ | `UI.gymMember` |
| B7 | 会計区分 | `UI.selfPay_accountingType` |
| B8 | 慢性候補フラグ | `UI.selfPay_chronicFlag` |
| D8 | 次回予約あり | `UI.selfPay_nextReserv` |
| F8 | 新規区分 | `UI.selfPay_firstVisitType` |
| A12:H13 | ケース1 部位・傷病・受傷日・治療法 | `UI.case1_rows` |
| A36:H37 | ケース2 部位・傷病・受傷日・治療法 | `UI.case2_rows` |
| A23:B28 | 所見（ケース1） | `UI.case1_shoken` |
| A16:B20 | 経過_今回（ケース1） | `UI.case1_keikaNow` |

Web UI から呼んだ場合、これらのセル値は「現在スプレッドシートに入力されている値」を読む。
Web App セッションとスプレッドシート操作セッションは別のため、
**Web から呼んでも意図した値が読めない。**

### `setPatientAndDate_V3(patientId)` の副作用

```javascript
uiSh.getRange(UI.patientDisplay).setValue(displayStr);  // B2 に表示文字列
uiSh.getRange(UI.treatDate).setValue(new Date());        // B4 に当日日付
```

Web から `setPatientAndDate_V3` を呼ぶと**シートの B2/B4 が書き換わる**。
スプレッドシートをもう一方の端末で開いている場合に干渉する可能性がある。

Phase WEB-2 の来院登録では、この関数を通じてシートに書き込む方式は採用しない。

---

## 2. `saveVisitFromWeb_V3` の方針

### 設計方針

- UI シートセルを**一切読まない**
- 引数（payload）からすべての入力値を受け取る
- 既存の `buildVisitKey_` を再利用する
- 既存の `saveVisit_V3` が行っているバリデーション・保存ロジックをできる限り再利用する
- 保険算定（`calcVisitAmounts_V3_` 等）は既存関数を呼ぶ
- Phase WEB-2 では部位・区分を中心にし、`shoken` / `keikaNow` は省略可（空欄許容）

### 関数シグネチャ

```javascript
/**
 * Web UI から来院を登録する（UI シート非依存版）。
 * @param {Object} payload
 * @param {string} payload.patientId
 * @param {string} payload.visitDate         - "YYYY-MM-DD"
 * @param {string} payload.accountingType    - "保険のみ"|"保険+自費"|"自費のみ"|""
 * @param {boolean} payload.gymMemberFlag
 * @param {boolean} payload.chronicCandidateFlag
 * @param {boolean} payload.nextReservation
 * @param {string} payload.firstVisitType    - "保険新規"|"自費直新規"|"再来"|""
 * @param {string} [payload.shoken]          - 所見（省略可）
 * @param {string} [payload.keikaNow]        - 経過_今回（省略可）
 * @param {Object[]} payload.cases           - 最大2件
 * @param {number}  payload.cases[].caseNo  - 1 or 2
 * @param {string}  payload.cases[].bodyPart
 * @param {string}  payload.cases[].disease
 * @param {string}  payload.cases[].injuryDate  - "YYYY-MM-DD"
 * @param {boolean} [payload.cases[].cold]
 * @param {boolean} [payload.cases[].warm]
 * @param {boolean} [payload.cases[].elec]
 * @param {string}  [payload.cases[].startDate] - "YYYY-MM-DD"
 * @param {string}  [payload.cases[].endDate]   - "YYYY-MM-DD"
 * @param {string}  [payload.cases[].tenki]
 * @returns {{ok:boolean, visitKey?:string, message?:string,
 *            reasonCode?:string, message?:string}}
 */
function saveVisitFromWeb_V3(payload) { ... }
```

### 返却仕様

**成功時:**
```javascript
{
  ok: true,
  visitKey: "PT001_2026-05-05",
  patientId: "PT001",
  message: "来院を登録しました"
}
```

**エラー時:**
```javascript
{
  ok: false,
  reasonCode: "DUPLICATE_VISIT",  // または MISSING_REQUIRED / INVALID_DATE / etc.
  message: "エラー内容（日本語）"
}
```

### reasonCode 一覧

| reasonCode | 意味 |
|---|---|
| `MISSING_REQUIRED` | 必須項目（部位・傷病・受傷日）が未入力 |
| `DUPLICATE_VISIT` | 同日同患者の二重登録 |
| `INVALID_DATE` | 日付形式不正 |
| `PATIENT_NOT_FOUND` | 患者IDが存在しない |
| `SYSTEM_ERROR` | GAS 内部エラー |

---

## 3. 保存先シートと列

### 来院ケースシート（`SHEETS.cases = "来院ケース"`）

来院ごとに 1〜2 行保存（caseNo=1 と caseNo=2）。

| CASE_COLS キー | 列名 | 役割 |
|---|---|---|
| `visitKey` | "visitKey" | 主キー（patientId_YYYY-MM-DD） |
| `treatDate` | "施術日" | 施術日 |
| `patientId` | "患者ID" | 患者ID |
| `caseNo` | "caseNo" | ケース番号（1 or 2） |
| `kubun` | "区分" | 算定区分（初検/再検/後療） |
| `injuryFixed` | "受傷日_確定" | 受傷日（確定版） |
| `p1` | "部位_部位1" | 第1部位 |
| `d1` | "傷病_部位1" | 傷病名 |
| `inj1` | "受傷日_部位1" | 受傷日 |
| `cold1` | "冷罨法_部位1" | 冷罨法フラグ |
| `warm1` | "温罨法_部位1" | 温罨法フラグ |
| `elec1` | "電療_部位1" | 電気療法フラグ |
| `shoken` | "所見" | 所見（省略可） |
| `keikaNow` | "経過_今回" | 経過_今回（省略可） |
| `caseKey` | "caseKey" | エピソードID（patientId_初検日_C{n}） |

### 来院ヘッダシート（`SHEETS.header = "来院ヘッダ"`）

来院ごとに 1 行保存。算定金額・KPI・監査列を含む。

| 列名 | 役割 |
|---|---|
| "visitKey" | 主キー |
| "施術日" | 施術日 |
| "患者ID" | 患者ID |
| "区分" | 算定区分 |
| "来院合計" | 算定合計金額 |
| "窓口負担額" | 窓口負担 |
| "保険請求額" | 保険請求額 |
| "会計区分" | 保険のみ/保険+自費/自費のみ |
| "ジム会員フラグ" | ジム会員フラグ |
| "要確認" | レセプト事故防止フラグ |
| "作成日時" | 保存日時 |

---

## 4. visitKey 発行方針

### 既存のルール（維持）

```javascript
function buildVisitKey_(patientId, treatDate) {
  return patientId + "_" + fmt_(treatDate, "yyyy-MM-dd");
}
```

- 形式: `{patientId}_{YYYY-MM-DD}`
- 例: `PT001_2026-05-05`
- 同日二重登録禁止: 来院ヘッダに同一 visitKey があれば `DUPLICATE_VISIT` エラー

**`buildVisitKey_` は `saveVisitFromWeb_V3` でそのまま再利用する。**

---

## 5. `getPrevVisitData_V3` の方針

### 目的

`autofillFromPreviousVisit_V3` はシートへの書き込み前提なので、
Web UI 用に「前回来院データを JSON で返す」版が必要。

### 関数シグネチャ

```javascript
/**
 * 指定患者の直前来院データを JSON で返す（読み取りのみ）。
 * @param {string} patientId
 * @returns {{ok:boolean, lastVisitDate?:string, cases?:Array, error?:string}}
 */
function getPrevVisitData_V3(patientId) { ... }
```

### 返却仕様

```javascript
{
  ok: true,
  patientId: "PT001",
  lastVisitDate: "2026-05-01",
  cases: [
    {
      caseNo: 1,
      bodyPart: "腰部",
      disease: "捻挫",
      injuryDate: "2026-04-01",
      cold: false,
      warm: true,
      elec: true,
      startDate: "2026-04-01",
      endDate: "",
      shoken: "...",
      keikaNow: "..."
    }
  ]
}
```

### 実装方針

1. `来院ケース` シートから指定患者の最新 `treatDate` を持つ行を取得
2. ケース1/ケース2 のデータを JSON に変換して返す
3. 個人情報はログに出力しない
4. `Logger.log` には `patientId` と件数のみ

---

## 6. 監査ログ方針

### 保存時のログ

`saveVisitFromWeb_V3` 実行時、以下のみ Logger に記録する:

```javascript
Logger.log("[saveVisitFromWeb_V3] action=WEB_VISIT_CREATE patientId=" + pid
  + " visitKey=" + visitKey + " result=" + (ok ? "OK" : reasonCode));
```

### ログ禁止フィールド

- 氏名・住所・電話番号・生年月日
- 保険者番号・記号番号・被保険者情報
- 部位名・傷病名（レセプト情報）

### 将来実装（Phase WEB-3 以降）

監査シートへの書き込みは Phase WEB-3 で検討する。
現時点では Logger のみで十分。

---

## 7. Phase WEB-2 実装手順

### 推奨実装順

```
Step 1: getPrevVisitData_V3(patientId) を実装
        → Ver3_core.js 末尾に追記（読み取りのみ、安全）

Step 2: web-visit-new.html を作成
        → 患者ID・来院日・会計区分・部位入力フォーム
        → 保存前確認モーダル
        → 「前回引き継ぎ」ボタン（getPrevVisitData_V3 呼び出し）

Step 3: doGet に page=visitNew を追加
        → patientId パラメータを受け取る

Step 4: web-patient-detail.html に「来院記録を追加」ボタン追加
        → ?page=visitNew&patientId=xxx へ遷移

Step 5: saveVisitFromWeb_V3(payload) を実装
        → 既存の saveVisit_V3 の処理をシート非依存で再実装
        → バリデーション → 来院ケース保存 → 来院ヘッダ保存 → 金額計算
        → 返却: {ok, visitKey, message}

Step 6: 保存後 web-patient-detail.html に戻り来院履歴が更新されることを確認

Step 7: clasp push

Step 8: 実機確認
        → 来院登録が来院ケースシートに反映されるか
        → 来院ヘッダに追加されるか
        → 既存の saveVisit_V3 による保存と競合しないか

Step 9: git commit + push
```

### Phase WEB-2 スコープ外（後回し）

以下は複雑すぎるため Phase WEB-2 には含めない:

| 機能 | 理由 | 候補フェーズ |
|---|---|---|
| 保険金額自動計算 | `calcVisitAmounts_V3_` が複数シートに依存 | WEB-2.5 |
| `keikaNow` / `shoken` 入力 | 複数セル結合・テキスト管理が複雑 | WEB-2.5 |
| 施術録生成 | Google ドキュメントのテンプレート操作 | WEB-3 |
| 申請書生成 | 帳票レイアウトが厚労省様式依存 | WEB-3 |
| 来院ヘッダへの金額記入 | 算定ロジック全体のテスト待ち | WEB-2.5 |

Phase WEB-2 のゴール:
**「部位・傷病・区分候補をWeb UIから入力して来院ケースに保存できる状態」**

---

## 8. web-home.html の home default 化条件

以下の条件が揃ったときに `doGet` のデフォルトを `home` に変更する:

1. `web-home.html` の実機確認 PASS
2. `patientSearch.html` に「← Web ホームへ」リンク追加済み（WEB-1B で実装）
3. `selfPayWeb.html` への既存導線（patientSearch → selfPayWeb）が壊れていない
4. `web-patient-detail.html` の実機確認 PASS
5. 患者詳細 → 来院記録追加の基本導線が安定（Phase WEB-2 完了後）
6. 現場でスマホ操作が試された実績

---

## 9. 未決定事項（Phase WEB-2 着手前に要確認）

| 項目 | 現状 | 決定が必要な理由 |
|---|---|---|
| `setPatientAndDate_V3` を Web 来院登録で使うか | 使わない方針 | 副作用がスプレッドシート運用に干渉 |
| `web-home.html` のデフォルト化時期 | 条件リストあり（§8） | 現場確認待ち |
| 来院保存後の金額計算 | 未実装 | `calcVisitAmounts_V3_` のシート依存度を調査要 |
| `keikaNow` / `shoken` の Web 入力 | Phase WEB-2.5 | セル結合の複雑さのため後回し |
