# CLAUDE.md

<!-- HIRAYAMA_WORKSPACE_WORKFLOW:START -->

## Hirayama Workspace 共通作業ルール

このrepoでの全作業にPersonal Skill
`hirayama-workspace-repo-workflow`
を使用する。

共通ルール正本：

`C:\hirayama-ai-workspace\CLAUDE_WORKFLOW_STANDARD.md`

ユーザーのプロンプトは今回の差分指示として扱い、
共通のGit確認、dirty停止、必須ファイル読込、Skills確認、
安全ルール、検証、記録、commit / push、最終報告を
自動適用する。

repo固有の情報は `README.md`、管理Markdown、
コード、Git履歴から確認する。

ユーザーへ共通ルールの再掲を求めない。

<!-- HIRAYAMA_WORKSPACE_WORKFLOW:END -->

## Repo固有ルール

現時点で未整理。
作業開始時に `README.md` と既存管理Markdownを確認し、
repo固有の制約が判明した場合は、この節へ記録する。

---

<!-- HIRAYAMA_BUSINESS_ENTITY_BOUNDARY_START -->
## 事業主体・屋号（中央正本を必ず読む）

| 項目 | 内容 |
|---|---|
| 中央正本 | `C:\hirayama-ai-workspace\BUSINESS_ENTITY_BOUNDARY_STANDARD.md` |
| repo 対応表 | `C:\hirayama-ai-workspace\BUSINESS_ENTITY_REPO_MAP.md` |

- **作業開始時に必ず中央正本を読む。** 区分の本文はここへ複製しない（正本を 2 つにしない）。
- **この repo の事業主体:** 接骨院事業＝**ひらやま接骨院（個人事業）**。対外表記に法人名を使わない。
- 主体が確定できない場合は、推測せず **`OWNER_CONFIRMATION_REQUIRED`** で停止し、owner 確認まで進めない。
- **法人事業（株式会社ひらやま）と個人事業（マシンやさんグループ／ワイルドボア／ひらやま接骨院）を推測で変更しない。**
  repo 名・ドメイン・メールアドレス・口座・既存文面を主体の根拠にしない。既存の正しい屋号表記を法人名へ統一しない。
<!-- HIRAYAMA_BUSINESS_ENTITY_BOUNDARY_END -->
