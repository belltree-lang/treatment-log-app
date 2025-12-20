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

## `resolveInvoiceReceiptDisplay_` / `formatAggregatedReceiptRemark_` 期待仕様（詳細）
空値/UNPAID/HOLD/AGGREGATE/合算無しの主要ケースをテーブル化し、表示可否と備考の期待値を明示する。

| receiptStatus 入力 | aggregateUntilMonth 入力 | 想定される整形結果 | showReceipt | receiptRemark | receiptMonths 例 | 備考 |
| --- | --- | --- | --- | --- | --- | --- |
| `null` / 空文字 | 未指定 | `receiptStatus: ''`（正規化）、`aggregateUntilMonth: ''` | `true` | `''` | `['202501']` | 「未指定は表示するが備考なし」の基準ケース |
| `UNPAID` | 未指定/指定 | `status: 'UNPAID'` | `false` | `''` | `['202501']` or `['202501','202502']` など | 支払待ちは常に非表示。合算指定があっても備考は付けない |
| `HOLD` | 未指定/指定 | `status: 'HOLD'` | `false` | `''` | `['202501']` or `['202501','202502']` など | 保留も非表示で固定 |
| `AGGREGATE` | `202503`（請求月 202501） | `status: 'AGGREGATE'`、`aggregateUntilMonth: '202503'` | `true` | `令和7年1月分・03月分施術代として` | `['202501','202502','202503']` | 月範囲を令和表記（開始は年付、以降は月のみ）で連結 |
| `AGGREGATE` | 未指定/空文字 | `status: 'AGGREGATE'`、`aggregateUntilMonth: ''` | `true` | `''` | `['202501']` | 集計終了月が無くても領収書は表示、備考は空 |
| `PAID` / その他 | 未指定 | `status: 'PAID'` | `false` | `''` | `['202501']` | 合算無しの通常ステータスは非表示 |
| `PAID` / その他 | `202503`（請求月 202501） | `status: 'PAID'`、`aggregateUntilMonth: '202503'` | `true` | `令和7年1月分・03月分施術代として` | `['202501','202502','202503']` | 合算指定があればステータスに関わらず表示し備考を付ける |

上記表は `resolveInvoiceReceiptDisplay_` と `formatAggregatedReceiptRemark_` の現行挙動に沿った期待値であり、今後のテストケース作成時の基準とする。【F:src/output/billingOutput.js†L284-L318】

## `aggregateUntilMonth` のバリデーション仕様（定義案）
- **フォーマット**: `YYYYMM` 6 桁数字。非数・桁不足は無効扱いとし、`normalizeInvoiceMonthKey_` で正規化できない値は空に落とす。【F:src/output/billingOutput.js†L246-L258】
- **範囲**: `200001`〜`209912` を有効とし、それ以外はエラー扱いにして空へフォールバック（上限/下限は Issue で調整可）。
- **請求月との前後関係**: `aggregateUntilMonth` は請求月以上を許容し、それより前なら無効とみなす。逆転時は単月に丸めず空へフォールバックする。
- **エラー時のフォールバック**:
  - `resolveInvoiceReceiptDisplay_` は請求月のみで `receiptMonths` を構成し、`showReceipt` は `receiptStatus` に基づく（`AGGREGATE` でも合算無しと同等）。
  - 備考は常に空文字。
  - UI もしくはログで「集計終了月が無効」のメッセージを出せるよう Issue を起票する。

## 合算ユーティリティの仕様適合確認と追加タスク
- `buildInclusiveMonthRange_` は開始 > 終了で開始のみ返却し、240 ヶ月で打ち切る現行挙動。上記バリデーションでは「請求月より前は無効にして空へフォールバック」にしたいため、逆転時の返却値を空にする修正が必要（Issue 化）。【F:src/output/billingOutput.js†L259-L282】
- `formatAggregatedReceiptRemark_` は月ラベル生成に失敗した要素を除外するため、部分的に無効な配列でも出力される。月範囲が欠落している場合は備考を空に統一する仕様としたいので、無効要素を含む場合は空文字を返すよう変更が必要（Issue 化）。【F:src/output/billingOutput.js†L284-L305】
- 上記 2 点を Issue に記載し、実装は後続とする。

## 受入条件（ロジック単体テスト入力例と期待出力）
- **空ステータス**: `{ billingMonth: '202501', receiptStatus: null, aggregateUntilMonth: null }` → `showReceipt: true`、`receiptRemark: ''`、`receiptMonths: ['202501']`。
- **UNPAID**: `{ billingMonth: '202501', receiptStatus: 'UNPAID', aggregateUntilMonth: '202503' }` → `showReceipt: false`、`receiptRemark: ''`、`receiptMonths: ['202501','202502','202503']`。
- **HOLD**: `{ billingMonth: '202501', receiptStatus: 'HOLD', aggregateUntilMonth: null }` → `showReceipt: false`、`receiptRemark: ''`、`receiptMonths: ['202501']`。
- **AGGREGATE + 正常月**: `{ billingMonth: '202501', receiptStatus: 'AGGREGATE', aggregateUntilMonth: '202503' }` → `showReceipt: true`、`receiptRemark: '令和7年1月分・03月分施術代として'`、`receiptMonths: ['202501','202502','202503']`。
- **AGGREGATE + 終了月欠落**: `{ billingMonth: '202501', receiptStatus: 'AGGREGATE', aggregateUntilMonth: '' }` → `showReceipt: true`、`receiptRemark: ''`、`receiptMonths: ['202501']`。
- **PAID + 合算無し**: `{ billingMonth: '202501', receiptStatus: 'PAID', aggregateUntilMonth: null }` → `showReceipt: false`、`receiptRemark: ''`、`receiptMonths: ['202501']`。
- **PAID + 合算指定**: `{ billingMonth: '202501', receiptStatus: 'PAID', aggregateUntilMonth: '202503' }` → `showReceipt: true`、`receiptRemark: '令和7年1月分・03月分施術代として'`、`receiptMonths: ['202501','202502','202503']`。
- **無効な aggregateUntilMonth**: `{ billingMonth: '202501', receiptStatus: 'AGGREGATE', aggregateUntilMonth: '20241234' }` → `aggregateUntilMonth` を無効として扱い、`showReceipt: true`、`receiptRemark: ''`、`receiptMonths: ['202501']`（バリデーション後のフォールバックを確認）。

## 請求月キー正規化と月範囲ビルダーの挙動
- `normalizeBillingMonthKeySafe_` は引数がオブジェクトの場合、`key` / `billingMonth` / `ym` / `month.key` / `month.ym` / `month` の順に候補を拾う。各候補を `normalizeBillingMonthInput` で厳密変換し、例外時は文字列化してトリムした値でフォールバックする。候補が全て空なら空文字を返す（空入力時の安全な戻り値）。【F:src/main.gs†L165-L189】
- 未来月や任意の 6 桁数字は `normalizeBillingMonthInput` が弾かないため、そのままキー化される（例: `209912` も通る）。月部分が 1〜12 以外なら例外が発生し、上記フォールバックでトリム済み文字列か空文字になる。【F:src/get/billingGet.js†L236-L271】
- `buildInclusiveMonthRange_` は開始キー/終了キーを双方 6 桁に正規化し、どちらか欠けた場合は開始キーがあれば単一要素、無ければ空配列で返す。開始 > 終了の逆順指定時は開始キーのみを返す。ループは 240 か月（20 年相当）で強制打ち切り、未来月を含んでいても上限に到達した時点で停止する。【F:src/output/billingOutput.js†L259-L282】

## `aggregateUntilMonth` のフォールバック仕様
- `updateBillingReceiptStatus` は `receiptStatus` が `AGGREGATE` のときのみ `aggregateUntilMonth` を正規化し、他のステータスでは空文字にリセットする。正規化に失敗すると空文字になり、以降の処理では「集計終了月なし」と同等に扱われる。【F:src/main.gs†L2017-L2035】【F:src/main.gs†L1455-L1493】
- `resolveInvoiceReceiptDisplay_` は `aggregateUntilMonth` が空や無効だった場合でも、請求月を基点に `receiptMonths` を最低 1 件返す（請求月のみ）。集計終了月が欠落したままでも領収書が表示され得るため、`AGGREGATE` で空に正規化されたケースに警告ログを入れるなら `mergeReceiptSettingsIntoPrepared_` かフロントの入力検証に追記する余地がある。【F:src/output/billingOutput.js†L298-L318】【F:src/main.gs†L433-L458】

## 合算対象月リストの上限ポリシー（Issue 追記案）
- `buildInclusiveMonthRange_` の while ループは 240 回で break するため、`receiptMonths` の最大長は 240。計算量を抑えつつ「20 年分まで表示すれば十分」という前提で組まれているとみられる。【F:src/output/billingOutput.js†L272-L281】
- 今後 Issue に記載する方針案: (1) 上限値 240 を仕様として明文化し、これを超える入力には警告ログを出す、(2) 240 を設定値として切り出し、長期請求に備える。パフォーマンス面の懸念が薄ければ (1) で十分だが、超長期データを扱う想定がある場合は (2) を検討する。

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

## 真偽の優先順位と整合性チェックシナリオ
- **真偽の優先順位**:
  - `savePreparedBillingToSheet_` は `PreparedBillingMeta` / `PreparedBillingMetaJson` / `PreparedBillingJson` の 3 シートへ「集計済みペイロード」を月単位で保存する。`preparedAt` は保存時刻に強制され、`preparedBy` も GAS の実行ユーザーが差し込まれるため、同月で複数回保存した場合はこちらが最新の正とみなされる。【F:src/main.gs†L491-L524】
  - `appendBillingHistoryRows` は履歴シートを月×患者 ID キーでマージ更新する。既存行があればレコードを上書き/補完するが、`receiptStatus` と `aggregateUntilMonth` は既存値を優先して保持する。【F:src/output/billingOutput.js†L1049-L1137】
  - 従って「集計結果の真実」は `savePreparedBillingToSheet_` 側にあり、履歴はそれを写経するが、一度書かれた領収関連セルはデフォルトでは履歴を手動で修正しない限り上書きされないという解釈になる。
- **整合性チェックシナリオ（Issue へ貼り付け可）**:
  - 同一月に対し `savePreparedBillingToSheet_` が複数回走ったとき、`appendBillingHistoryRows` が最新の `receiptStatus` / `aggregateUntilMonth` を反映できているか。
  - 履歴シートに手動変更が入った場合（例: `receiptStatus` を直接編集）、次回の `appendBillingHistoryRows` で準備済みデータと食い違いが起きていないか。
  - `billingMonth` と `patientId` キーが一致しないゴミ行がある場合に、`appendBillingHistoryRows` が意図せず残存データを参照しないか。
  - `PreparedBillingJson` の行削除（該当患者が請求対象から外れたケース）後に `appendBillingHistoryRows` が古い履歴行をクリアできているか（`clearContent` で上書きされる想定だが、キーが欠けた行が残らないか）。

## 既存値優先ロジックで領収状態を上書きできない懸念
- `appendBillingHistoryRows` は `receiptStatus` / `aggregateUntilMonth` に限り「既存セルが空でない場合はそのまま残す」ロジック。つまり `UNPAID→AGGREGATE` や `HOLD→''` のような変更が準備済みデータ側に入っても、履歴シートでは反映されず保持される。【F:src/output/billingOutput.js†L1109-L1133】
- 以下の確認項目をテストや手動チェックに追加する:
  - 既存履歴に `receiptStatus: UNPAID` が入った状態で `billingJson` が空文字ステータスを返した場合、履歴は更新されず前回値が残ることを確認する（仕様なら OK、上書きしたいならロジック変更を検討）。
  - `aggregateUntilMonth` を `202503→''` に戻したケースで、履歴が旧値を保持していないか。
  - 履歴行が存在しない新規患者では、準備済みデータの `receiptStatus` / `aggregateUntilMonth` がそのまま入ることを確認する。

## 監査用メタデータの要否
- `savePreparedBillingToSheet_` は `preparedAt` / `preparedBy` / `schemaVersion` をメタシートへ保存するが、`appendBillingHistoryRows` 側で「誰が履歴を書き換えたか」は `updatedAt`（現在時刻）しか残らない。【F:src/main.gs†L491-L520】【F:src/output/billingOutput.js†L1019-L1046】【F:src/output/billingOutput.js†L1088-L1113】
- 履歴シートが監査対象なら、以下の差分記録を Issue に追記する余地がある:
  - `updatedBy`（最終更新者メール）列を追加し、GAS 実行ユーザーを保存する。
  - `updatedReason`（自動反映 / 手動追記 / 再計算 など）をオプションで受け取り、`appendBillingHistoryRows` で入れられるようにする。
  - `updatedAt` の更新タイミングを「既存行と差分があったときのみ」に絞り、差分内容（変更フィールドと旧値/新値）を `PreparedBillingMetaJson` の note 欄や別シートでログ化する。
