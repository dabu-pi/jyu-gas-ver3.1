# WEB-6: 共通グローバルナビタブ追加記録

作成日: 2026-05-08  
ステータス: **LiveCheck 9 PASS 確認済み / 本番 deploy 可（2026-05-08）**

---

## 概要

各 Web ページの上部に、共通のグローバルナビゲーションタブ（`.web-nav`）を追加した。  
どのページからでも HOME / 患者検索 / 月次申請 / 自費明細へ移動できる。

---

## 更新したページ一覧

| ページファイル | page= | active タブ | 旧ナビ削除 |
|---|---|---|---|
| `web-home.html` | home | 🏠 HOME | — |
| `patientSearch.html` | search | 患者検索 | ← Web ホームへ リンク削除 |
| `web-patient-detail.html` | detail | 患者検索 | — |
| `web-visit-new.html` | visitNew | 患者検索 | — |
| `web-monthly-claims.html` | monthlyClaims | 月次申請 | ← Web ホームへ リンク削除 |
| `web-monthly-claim-detail.html` | monthlyClaimDetail | 月次申請 | — |
| `selfPayWeb.html` | selfpay | 自費明細 | — |

---

## ナビリンク構成

```html
<nav class="web-nav">
  <a href="<?= appBaseUrl ?>?page=home" target="_top">🏠 HOME</a>
  <a href="<?= appBaseUrl ?>?page=search" target="_top">患者検索</a>
  <a href="<?= appBaseUrl ?>?page=monthlyClaims" target="_top">月次申請</a>
  <a href="<?= appBaseUrl ?>?page=selfpay" target="_top">自費明細</a>
</nav>
```

各ページでは対応するタブに `class="active"` を付与。

---

## iframe 白画面対策

**全ナビリンクに `target="_top"` を付与。**

GAS Web App は2段 iframe 構造のため、`target="_top"` なしのリンクをクリックすると inner iframe 内で遷移し、GAS 構造が入れ子になって白画面になる（過去に発生した既知バグ）。

---

## CSS（全ページ共通）

```css
.web-nav{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px;padding:8px 10px;
  background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;font-size:13px}
.web-nav a{display:inline-block;padding:6px 12px;border-radius:20px;
  text-decoration:none;font-weight:600;color:#334155;background:#fff;
  border:1px solid #cbd5e1;white-space:nowrap}
.web-nav a:hover{background:#f1f5f9}
.web-nav a.active{color:#fff;background:#2563eb;border-color:#2563eb}
```

スマホ幅では `flex-wrap:wrap` で折り返す。

---

## LiveCheck（WEB-6）— auth 更新後実行結果

| テスト | 内容 | 結果 |
|---|---|---|
| W6-1 | page=home に .web-nav が表示される（HOME active） | ✅ PASS |
| W6-2 | page=search に .web-nav が表示される（患者検索 active） | ✅ PASS |
| W6-3 | page=monthlyClaims に .web-nav が表示される（月次申請 active） | ✅ PASS |
| W6-4 | page=monthlyClaimDetail に .web-nav が表示される（月次申請 active） | ✅ PASS |
| W6-5 | 全ナビリンクに target="_top" が付いている | ✅ PASS |
| W6-6 | HOMEリンクの href が ?page=home を含む | ✅ PASS |
| W6-7 | 月次申請リンクの href が ?page=monthlyClaims を含む | ✅ PASS |
| W6-8 | monthlyClaimDetail でB案Excel生成ボタンが維持されている | ✅ PASS |
| W6-9 | iframe 入れ子増殖がない（白画面バグ再発なし） | ✅ PASS |

**9 PASS / 0 FAIL / 0 SKIP**

## dev 目視確認結果

| ページ | nav表示 | active | links | target | iframe段数 |
|---|---|---|---|---|---|
| home | ✅ | 🏠 HOME | 4 | ✅ _top | ✅ 3段 |
| search | ✅ | 患者検索 | 4 | ✅ _top | ✅ 3段 |
| detail | ✅ | 患者検索 | 4 | ✅ _top | ✅ 3段 |
| visitNew | ✅ | 患者検索 | 4 | ✅ _top | ✅ 3段 |
| monthlyClaims | ✅ | 月次申請 | 4 | ✅ _top | ✅ 3段 |
| monthlyClaimDetail | ✅ | 月次申請 | 4 | ✅ _top | ✅ 3段 |
| selfpay | ✅ | 自費明細 | 4 | ✅ _top | ✅ 3段 |

**全7ページ nav 表示 / active / target="_top" / iframe 非増殖 確認済み**

---

## clasp push

実施済み（2026-05-08 17:21:03 / 20ファイル）

---

## 本番 deploy

**未実施**  
dev URL で目視確認後に deploy を判断する。

**dev URL 確認用リンク:**
```
?page=home
?page=search
?page=detail&patientId=hirayamaka
?page=visitNew&patientId=hirayamaka
?page=monthlyClaims
?page=monthlyClaimDetail&patientId=hirayamaka&ym=2026-04
?page=selfpay
```

---

## B案ルートへの影響

| 対象 | 変更 |
|---|---|
| `generateClaimApplicationBFromWeb_V3` | **変更なし** |
| `V3TR_menuGenerateApplication_B` | **変更なし** |
| `getMonthlyClaimDetail_V3` / `buildMonthlyTransferData_V3` | **変更なし** |

---

## 残確認事項（人間目視）

- [ ] dev URL で各ページのナビが表示されること
- [ ] active タブが正しいページで強調されること
- [ ] スマホ幅でナビが折り返して読めること
- [ ] どのページからでも HOME に戻れること
- [ ] 月次申請詳細でナビを使っても白画面にならないこと

---

## Dashboard

**対象外**
