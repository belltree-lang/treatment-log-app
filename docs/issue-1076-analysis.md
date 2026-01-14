# Issue 1076 調査メモ（receipt-debug ログ前提）

## 背景
AE/AF 未設定でも前月分領収書が出る経路を特定するため、`[receipt-debug]` ログを `receiptMonths` 設定直前に追加し、患者 ID を限定できるガード（`BILLING_DEBUG_PID`）を用意しました。ログには `currentFlags` / `previousFlags` と `receiptTargetMonths` / `receiptMonths` を出力する想定です。

## ログから読み取れる注入経路（想定）
`resolveReceiptTargetMonthsFromBankFlags_` の判定は以下のため、**現月・前月とも AE/AF が未設定**の場合に「前月のみ」が `receiptTargetMonths` として返されます。

1. 現月に `ae`/`af` がある場合は `receiptTargetMonths = []`。
2. 前月に `af` がある場合は、未収合算対象月（`collectAggregateBankFlagMonthsForPatient_`）+ 前月を `receiptTargetMonths` にします。
3. 前月に `ae` がある場合は `receiptTargetMonths = []`。
4. 上記いずれでもない場合は **`receiptTargetMonths = [previousMonthKey]`**。

このため **AE/AF 未設定でも前月が入る**主要経路は、`resolveReceiptTargetMonthsFromBankFlags_` の「デフォルト分岐（4）」であると結論づけられます。ここで生成された `receiptTargetMonths` が `attachPreviousReceiptAmounts_` によって `receiptMonths` としてセットされ、最終的に `finalizeInvoiceAmountDataForPdf_` の `receiptMonths` に伝播します。
