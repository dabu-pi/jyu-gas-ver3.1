# JYU-GAS Source of Truth 明確化 — 2026-05-14

## 背景

2026-05-14 の workspace git health check で、`gas-projects/jyu-gas-ver3.1` の以下 6 ファイルが
**HEAD（= origin/main）には tracked、ディスクには欠損**という状態で検出された。

| ファイル | HEAD bytes | 役割 |
|---|---|---|
| `Ver3_core.js` | 300,004 | 来院登録・区分判定・算定ロジック |
| `Ver3_amounts.js` | 54,197 | 金額計算（初検料・再検料・逓減・長期減額） |
| `Ver3_transferData.js` | 129,354 | 申請書データ転記（B 案メニュー本体） |
| `Ver3_patientPicker.js` | 6,906 | 患者選択ダイアログ |
| `SPEC.md` | 20,921 | 金額計算仕様書 |
| `appsscript.json` | 244 | clasp deploy 必須のマニフェスト |

これらは現行 disk の `Ver3_test.js` / `Ver3_shuRecorder.js` から実コード上で関数呼び出し参照されている
active production ファイルである（例: `calcOnePartAmount_V3_`、`V3TR_buildTransferDataForMonth_`）。

詳細経緯は `../../docs/GIT_DIRTY_ROOT_CAUSE_2026-05-14.md` を参照。

---

## 危険度

- **HIGH**: この状態のまま `clasp push` を実行すると、GAS 上の production code（@13 deploy 済）を削除する可能性がある
- 患者保険請求 production system のため、データ損失リスクあり

---

## Source of Truth ルール（確定）

| 項目 | 正本 |
|---|---|
| コード本体 | GitHub `dabu-pi/jyu-gas-ver3.1` の `main` branch |
| 仕様文書（SPEC.md / PROJECT_STATUS.md / docs/）| 同 GitHub |
| GAS deployment（`@N`）| clasp deploy で同期する派生物。**正本ではない** |
| `.claspignore` で除外されるもの | `SPEC.md` / `TESTCASES.md` / `PLAN.md` / `*.py` / `docs/**` / `*.pdf` 等 |
| `appsscript.json` | clasp 必須 → git 管理下に置く（除外しない） |

GAS editor は deploy 手段であって正本ではない。ローカル disk と GitHub が乖離したら GitHub を信用する。

---

## 今回採用した対応

1. `git checkout -- SPEC.md Ver3_amounts.js Ver3_core.js Ver3_patientPicker.js Ver3_transferData.js appsscript.json` で HEAD から復元
2. `git update-index --refresh` 後 `git status --porcelain=v1` 空・`git ls-files -d` 空を確認
3. 本書を `docs/` に追加して経緯と再発防止策を記録

---

## 再発防止ルール

### A. clasp push / clasp deploy 前の必須チェック

実行前に必ず以下を行い、1 件でも該当があれば push を停止する。

```powershell
cd C:\hirayama-ai-workspace\workspace\gas-projects\jyu-gas-ver3.1
git update-index -q --refresh
git ls-files -d
```

`git ls-files -d` が空でない場合、disk から欠損したまま push すると GAS 上の対応ファイルが削除される。

### B. ファイル削除は必ず git rm + commit まで

任意のファイルを意図的に削除する場合は同セッション内で:

```powershell
git rm <file>
git commit -m "remove(jyu-gas): <reason>"
git push
```

ディスクから消すだけで放置しない。

### C. 健全性チェックは workspace 共通スクリプトで

`C:\hirayama-ai-workspace\workspace\tools\git-health-check.ps1` を sync 前後に実行する。
missing tracked が検出されたらこの SOURCE_OF_TRUTH ルールに従って復元 or 削除する。

---

## 関連ドキュメント

- `../../docs/GIT_DIRTY_ROOT_CAUSE_2026-05-14.md` — workspace 全体の root cause 分析
- `../../tools/git-health-check.ps1` — sync 前後の必須監査スクリプト
- `PROJECT_STATUS.md` — JYU-GAS 全体ステータス
