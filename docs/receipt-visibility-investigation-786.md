# 領収書非表示の原因調査（issue #786）

## showReceipt / previousReceipt.visible が `false` になる経路
- **前月シート有無フラグが `false` のままになる場合**  
  - `attachPreviousReceiptAmounts_` が `billingMonth` 未設定・不正（数値化できないなど）の場合は早期 return し、`hasPreviousReceiptSheet` / `hasPreviousPrepared` を付与しない。【F:src/main.gs†L548-L576】  
  - `resolvePreviousBillingMonthKey_` が空文字を返す（例: `billingMonth` が `YYYY-MM` 以外で正規化できない）と、銀行引落シート検索を行わずフラグを `false` のまま返す。【F:src/main.gs†L548-L555】【F:src/output/billingOutput.js†L325-L335】  
  - `formatBankWithdrawalSheetName_` が生成するシート名と実シート名が合わない場合（`normalizeBillingMonthInput` で `2025/01` → `2025-01` にならないなど）も `getSheetByName` 失敗で `hasSheet: false` となる。【F:src/main.gs†L579-L619】【F:src/main.gs†L1401-L1418】
- **請求月の正規化に失敗して前月キーが作れない場合**  
  - `resolvePreviousBillingMonthKey_` は `normalizeInvoiceMonthKey_` に失敗すると空文字を返し、そのまま `receiptMonths` が空となり `showReceipt` が `false` になる。【F:src/output/billingOutput.js†L309-L335】
- **テンプレート描画時にフラグが再び上書きされる場合**  
  - `buildInvoiceTemplateData_` は `previousReceipt.visible` が `undefined` のとき `showReceipt` を踏襲するが、`hasPreviousReceiptSheet` が `false` だと最終的に強制で `false` にする。前段でフラグが欠落しているとここで非表示に固定される。【F:src/output/billingOutput.js†L438-L485】  
  - HTML テンプレートは `receipt.visible` → `receiptVisible` → `showReceipt` の優先順で真偽を決め、いずれも falsy なら `false` になる。【F:src/invoice_template.html†L85-L106】

上記のため、銀行引落シート（`銀行引落_YYYY-MM`）が存在していても、**請求月の正規化失敗や前月シート有無フラグが未設定のまま**だと表示判定が `false` に落ちる。

## 現行思想で不要と判断できる領収書関連コード
- `isPreviousReceiptSettled_` は常に `true` を返すだけで利用箇所もなく、設計ドキュメントでも採用しないと明記されているため削除候補。【F:src/output/billingOutput.js†L338-L340】【F:docs/billing-receipt-status.md†L190-L201】
- `receiptVisible` / `data.showReceipt` などテンプレートの多段優先ロジックは、実質 `previousReceipt.visible` と同義になっており、設計方針（未回収チェックのみで判定）と重複するため簡素化余地あり。【F:src/output/billingOutput.js†L438-L485】【F:src/invoice_template.html†L85-L106】
- 銀行シート由来の前月領収金額付与ロジック（`attachPreviousReceiptAmounts_` や `previousReceiptAmount` 依存の付帯データ）は、設計上「銀行結果を判定に使わない」とされており縮小対象。【F:src/main.gs†L548-L576】【F:docs/billing-receipt-status.md†L182-L203】
