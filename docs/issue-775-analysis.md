# Issue 775 削除対象ロジックの静的参照状況

本ドキュメントは Issue #775 の削除候補ロジックについて、現行コードの参照・依存関係を静的に確認した結果をまとめたものです。テスト実行は行っていません。

## 分析観点
- 参照元・依存関係の列挙
- 削除時に影響する機能（銀行引落CSV生成、当月請求計算、新規患者の初回請求）
- 削除可否の判定（完全削除可 / 呼び出し元修正が必要 / 削除不可）

## 判定サマリ

| 対象 | 主要な参照箇所 | 機能依存 | 判定 |
| --- | --- | --- | --- |
| `loadPreparedBillingWithSheetFallback_` / `normalizePreparedBilling_` | 請求キャッシュ読込、請求書生成、銀行引落シート生成・確認で広範に使用【F:src/main.gs†L1052-L1092】【F:src/main.gs†L1930-L1975】 | 請求集計結果の読込ができなくなるため請求PDF生成・銀行引落シート生成が崩れる | 削除不可 |
| `buildBillingAmountByPatientId_` | 銀行引落シートへの金額書込み、前月領収金額付与、金額差分ログで使用【F:src/main.gs†L1260-L1314】【F:src/main.gs†L596-L610】【F:src/main.gs†L712-L729】 | 金額列を埋められず銀行引落CSV生成と請求書の前月金額表示に支障 | 削除不可 |
| `resolvePreviousBillingMonthKey_` | 未回収履歴走査・前月領収金額取得で使用【F:src/main.gs†L460-L503】【F:src/main.gs†L583-L610】 | 前月キー解決ができず未回収履歴判定・前月領収金額付与が停止 | 削除不可 |
| `applyReceiptRulesFromUnpaidCheck_` 内の未回収履歴走査・自動合算付与 | 請求書生成時に未回収チェック履歴から `unpaidChecked`/`aggregateUntilMonth` を付与【F:src/main.gs†L525-L581】【F:src/main.gs†L1958-L1975】 | 未回収チェック済み患者でも領収書が表示される等、請求書ロジックが変化 | 呼び出し元修正が必要（現仕様に合わせた別判定へ置換が必要） |
| `previousReceiptAmount` 付与ロジック（`attachPreviousReceiptAmounts_`） | 前月 prepared を読み前月領収金額と表示可否を設定【F:src/main.gs†L583-L610】【F:src/output/billingOutput.js†L335-L352】【F:src/output/billingOutput.js†L466-L488】 | 請求書の前月領収額表示と `hasPreviousPrepared` に依存する可視制御が消える。機能維持には代替データ供給が必要 | 呼び出し元修正が必要 |
| `isPreviousReceiptSettled_` | 定義のみで実質未使用【F:src/output/billingOutput.js†L331-L344】 | 既存機能に影響なし | 完全に削除可能 |
| `bankStatus` / `paidStatus` 依存 | 請求JSON生成で状態を埋め、銀行振込データ出力で列を更新【F:src/logic/billingLogic.js†L37-L94】【F:src/output/billingOutput.js†L656-L794】【F:src/output/billingOutput.js†L851-L888】 | 銀行振込（旧仕様）シートの「領収状態」列が空になる。状態情報を不要とする仕様変更と合わせて呼び出し側修正が必要 | 呼び出し元修正が必要 |
| 銀行引落金額と請求金額の差分ログ出力（`logBankWithdrawalAmountMismatches_`） | 銀行引落金額収集時に警告ログを出すのみで他ロジックへは未連携【F:src/main.gs†L652-L699】【F:src/main.gs†L701-L730】 | 出力データは変わらず、ログのみ抑止 | 完全に削除可能 |

## 詳細観察
- **銀行引落CSV生成への影響**: `generateSimpleBankSheet` が `buildBillingAmountByPatientId_` を使って銀行引落シートの金額列を埋めるため、同関数を削除すると金額未設定となる【F:src/main.gs†L1260-L1314】。
- **当月請求金額計算への影響**: 請求書生成フロー `generatePreparedInvoices_` で `normalizePreparedBilling_` → `applyReceiptRulesFromUnpaidCheck_` → `attachPreviousReceiptAmounts_` を順に適用しているため、いずれかを除去すると現在の請求書データ構成が変わる【F:src/main.gs†L1958-L1975】。
- **新規患者の初回請求**: 請求JSON生成時に `bankStatus`/`paidStatus` などを含めているが、金額計算自体は `generateBillingJsonFromSource` の計算ロジックに依存しており、候補削除に直結する箇所は確認されなかった【F:src/logic/billingLogic.js†L37-L94】【F:src/logic/billingLogic.js†L240-L302】。ただし前月データ付与 (`attachPreviousReceiptAmounts_`) を削除する場合、新規患者（前月データなし）では `hasPreviousPrepared` フラグ挙動が変わり、請求書テンプレートの前月領収欄表示条件が変化する点に注意【F:src/output/billingOutput.js†L466-L488】。
