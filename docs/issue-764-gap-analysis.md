# Issue #764 差分確認メモ

## 依頼内容
- 「前月領収書の発行判定は未回収チェックのみ」とする確定設計（docs/billing-receipt-status.md）と現行実装の差分を事実ベースで確認。
- 追加で判明した設計との差分・暗黙依存を列挙。

## 設計で明示されている方針
- 前月領収書の発行可否は、PDF発行時点の未回収チェックだけで決定する。他条件は持たない。【F:docs/billing-receipt-status.md†L151-L162】
- 前月領収書の金額は前月請求金額そのものを用い、銀行結果・previousReceiptAmount などは参照しない。【F:docs/billing-receipt-status.md†L163-L178】
- 銀行情報はCSV生成や履歴保存のみで使い、PDF発行可否判定には一切使わない。【F:docs/billing-receipt-status.md†L178-L188】
- settled/unsettled 判定や isPreviousReceiptSettled_、previousReceiptAmount、bankStatus による表示制御は採用しない。【F:docs/billing-receipt-status.md†L190-L203】

## 現行実装で確認できる相違
- 請求書PDF生成は常に `applyReceiptRulesFromUnpaidCheck_` を通し、銀行引落シートの未回収チェック履歴から `receiptStatus`/`aggregateUntilMonth` を自動付与する。未回収チェックがある患者は `HOLD` となり、連続未回収なら自動で合算指定も付く。【F:src/main.gs†L525-L587】【F:src/main.gs†L1969-L1993】
- 前月領収書金額や内訳は `collectBankWithdrawalAmountsByPatient_` で銀行引落シートから拾った金額を基に埋められ、未回収チェック付き行はスキップされる。previousReceiptAmount が無ければ 0 になり、銀行データ取得が前提になっている。【F:src/main.gs†L595-L712】
- 領収書表示判定は `receiptStatus` や `aggregateUntilMonth` に依存しており、`HOLD`/`UNPAID` なら非表示、合算指定があれば備考付きで表示といった旧ロジックが残っている。【F:src/output/billingOutput.js†L309-L364】
- 前月領収書自体の表示可否は `previousReceiptAmount` が正数かどうか（`isPreviousReceiptSettled_`）で上書きされ、0 や空の場合は自動的に非表示になるため、「未回収チェックのみで決定する」設計とは異なる。【F:src/output/billingOutput.js†L367-L378】【F:src/output/billingOutput.js†L476-L523】
- invoice テンプレートも `previousReceipt.visible` や `receiptVisible` などプログラム側の判定結果に従って表示/非表示が切り替わる構造で、未回収チェックだけでは制御されない。【F:src/invoice_template.html†L83-L136】

## 追加で見つかった暗黙依存・前提
- 未回収チェック判定は「銀行引落シートに未回収チェック列が存在し、`summarizeBankWithdrawalSheet_` などが利用可能である」ことを前提としており、シートが無い場合は未回収反映がスキップされる。【F:src/main.gs†L525-L540】【F:src/main.gs†L665-L712】【F:src/main.gs†L756-L780】
- 前月領収書金額や内訳は銀行引落シートの金額と患者紐付け（患者ID列または氏名照合）に依存するため、前月請求額の保存や再利用は行われていない。【F:src/main.gs†L665-L712】
- 請求データ生成時に銀行金額と請求計算額の不一致を監視するロガーが走り、銀行金額を集計できる前提で動作する。【F:src/main.gs†L714-L754】
