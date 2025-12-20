# 領収書表示ロジックとテスト計画

## `resolveInvoiceReceiptDisplay_` の組み合わせ表
`receiptStatus`・`aggregateUntilMonth`・請求月の関係で返却値がどう変わるかを整理しました。`receiptMonths` は `billingMonth` と `aggregateUntilMonth` の範囲を `buildInclusiveMonthRange_` で決定し、`aggregateUntilMonth` が無い場合は請求月のみになります。

| receiptStatus | aggregateUntilMonth | 請求月 (`billingMonth`) | showReceipt | receiptRemark | receiptMonths の例 |
| --- | --- | --- | --- | --- | --- |
| `null` / 空文字 | 未指定 | `202501` | `true` | 空文字 | `["202501"]` |
| `AGGREGATE` | `202503` | `202501` | `true` | `令和7年1月分・03月分施術代として`（`formatAggregatedReceiptRemark_` により範囲表記） | `["202501","202502","202503"]` |
| `AGGREGATE` | 未指定 | `202501` | `true` | 空文字 | `["202501"]` |
| `UNPAID` または `HOLD` | 任意 | `202501`（例） | `false` | 空文字 | `aggregateUntilMonth` があれば範囲、無ければ `["202501"]` |
| 上記以外（例: `PAID`） | 未指定 | `202501` | `false` | 空文字 | `["202501"]` |
| 上記以外（例: `PAID`） | `202503` | `202501` | `true`（`aggregateUntilMonth` が指定されていれば強制表示） | `令和7年1月分・03月分施術代として` | `["202501","202502","202503"]` |

- `UNPAID` / `HOLD` は常に非表示（`showReceipt: false`）。その他のステータスは `null` / 空文字 / `AGGREGATE` / `aggregateUntilMonth` 指定時のみ `true` になる。【F:src/output/billingOutput.js†L298-L318】
- `aggregateUntilMonth` がある場合のみ備考（`receiptRemark`）が付き、請求月から指定月までを令和表記で連結する。【F:src/output/billingOutput.js†L284-L318】

## 領収状態保存の前提（フロント → Apps Script → 履歴）
1. フロントエンドでステータス／集計終了月が変更されると `handleReceiptStatusChange` / `handleReceiptAggregateChange` が `billingState` を更新し、`persistReceiptStatus` を呼び出す。【F:src/main.js.html†L211-L256】
2. `persistReceiptStatus` は請求月未選択をブロックし、`AGGREGATE` 以外では `aggregateUntilMonth` を空に初期化した上で Apps Script の `updateBillingReceiptStatus` を実行。成功時にレスポンスでフロントの状態を再同期する。【F:src/main.js.html†L227-L256】
3. サーバー側の `updateBillingReceiptStatus` はステータスを正規化し、`mergeReceiptSettingsIntoPrepared_` で集計済みペイロードに反映する。結果をキャッシュ（`savePreparedBilling_`）し、同時にスプレッドシートのメタ・JSONシートへ保存（`savePreparedBillingToSheet_`）して履歴を残す。【F:src/main.gs†L433-L458】【F:src/main.gs†L772-L813】【F:src/main.gs†L2031-L2047】

## 単体テスト方針と必要データ（Issue 下書き）
- **テストランナー方針**: 既存の `tests/billingOutput.test.js` と同様に Node + `vm` で `billingOutput.js` を読み込み、Apps Script API はスタブに置き換える。`formatMonthWithReiwaEra_` など同ファイル定義はそのまま利用する。【F:tests/billingOutput.test.js†L1-L75】
- **モックデータ形式**:
  - 請求行オブジェクトは `{ billingMonth, receiptStatus, aggregateUntilMonth }` を基本とし、月の比較が分かるよう `billingMonth: '202501'` 固定で `aggregateUntilMonth` を `null` / `'202503'` などで切り替える。
  - 備考生成確認用に `formatMonthWithReiwaEra_` の出力へ依存するため、`aggregateUntilMonth` を跨ぐ配列（例: `['202501','202502','202503']`）を期待値として持つ。
- **追加で用意するテストケース**（想定 Issue チェックリスト）:
  - [ ] `receiptStatus` が `null` / 空文字の場合に `showReceipt: true`・備考なし・請求月のみになる。
  - [ ] `UNPAID` / `HOLD` が指定された場合に `showReceipt: false` かつ集計月配列のみ返る。
  - [ ] `AGGREGATE` かつ `aggregateUntilMonth` 指定時に、請求月から終了月までの `receiptMonths` と令和表記備考が付与される。
  - [ ] `AGGREGATE` だが `aggregateUntilMonth` が無い場合に備考が空のままになる。
  - [ ] `PAID` など未知ステータスでも `aggregateUntilMonth` があれば `showReceipt: true` になり、無ければ `false` になる。
- **Apps Script 互換層**: 本関数は GAS 固有 API を使わないため、GS スタブは不要。将来的に `formatAggregatedReceiptRemark_` のロケール依存が問題化した際は `Session.getScriptTimeZone` などのスタブを追加する余地がある。

以上を Issue に転記すれば、領収書表示ロジックの仕様とテストカバレッジ要求を共有できます。
