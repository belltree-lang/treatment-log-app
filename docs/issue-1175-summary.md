# Issue 1175 共有サマリー

## 変更点
- 請求PDF生成時に `applyEdits` を送らないようにし、PDF生成が編集保存を要求しないように変更。
- サーバー側でも `generatePreparedInvoicesForMonth` が `applyEdits` を受け取っても無視するようにし、PDF生成時に `applyBillingEdits` が実行されないように変更。

## 変更理由
- PDF生成フローがPreparedBillingや銀行引落シートに書き込みを行うと、AE/AF/AGフラグ（合算など）が意図せず解除・再計算される可能性があるため、完全に参照専用に固定する。

## 影響範囲
- 影響対象: 請求PDF生成（PreparedBillingを選択してPDF生成する導線）。
- 影響外: 集計処理、手動保存（請求編集の保存）、請求金額算出ロジック、PDFテンプレート。
