# 請求書／領収書コード調査メモ（issue #934）

## 概要
- 対象: `src/output/billingOutput.js` の領収・合算関連ロジック、および `src/main.js.html` の請求UI表示条件。
- 観点: 未使用コード／重複・責務不明な分岐、現行仕様で必須に見える表示条件の整理。挙動変更は実施していません。

## GS (billingOutput.js)

### 削除候補
- `buildAggregateInvoiceDetailForMonth_` は定義のみで呼び出し箇所が見当たらず、現在のテンプレート生成フローから外れている。履歴表示などで未使用。 【F:src/output/billingOutput.js†L757-L790】
- `formatInvoiceChargeMonthLabel_` は請求／領収書のどの出力パスからも参照されていない。 【F:src/output/billingOutput.js†L609-L618】

### 要整理・責務不明
- `resolveInvoiceReceiptDisplay_` のデフォルトでは `hasPreviousReceiptSheet` が未定義でも true となり、`receiptStatus !== 'UNPAID'` かつ銀行フラグがない限り領収書表示を許可する挙動。前月領収書の有無が未設定でも表示されるため、フラグ未設定と実在の区別がつかない。 【F:src/output/billingOutput.js†L362-L435】
- `resolveAggregateInvoiceDecision_` で計算する `fallbackReceiptMonths` は決定ロジックに使われず、`fallbackReceiptMonthsUsedInDecision` も常に false のまま trace に残るだけ。過去領収月を合算判定に組み込む意図が読み取れず、実際の意思決定と差異がある可能性。 【F:src/output/billingOutput.js†L437-L483】
- 同関数では `previousReceiptAmount` が 0 超なら合算対象月が 1 か月でも `isAggregateInvoice` が true になるため、領収書金額だけで合算モードになるケースがある。金額を持つだけで合算扱いにする意図か要確認。 【F:src/output/billingOutput.js†L437-L483】
- `resolveReceiptMonthBreakdown_` は `receiptMonthBreakdown` プロパティが無ければ外部の `buildReceiptMonthBreakdownForEntry_` に依存しており、同ファイル内に定義がない。責務が別モジュールに分かれているが所在が明示されず、実行環境により空配列になる可能性がある。 【F:src/output/billingOutput.js†L579-L592】
- `resolveHasPreviousReceiptSheet_` は `hasPreviousReceiptSheet`/`hasPreviousPrepared` が無い場合に true を返すため、未入力を「有り」と解釈する。領収書有無の実データと未設定の区別を付けたい場合は見直しが必要。 【F:src/output/billingOutput.js†L362-L370】

### 現行で必須のロジック
- `buildInvoiceTemplateData_`/`buildAggregateInvoiceTemplateData_` は領収月の決定、前月領収書表示の可否（`hasPreviousReceiptSheet` との AND）、合算決定トレースの生成を行う中核。請求 PDF 出力がこれらの戻り値を直接テンプレートに埋め込むため、領収表示条件と合算判定の整合性を保つ重要な経路。 【F:src/output/billingOutput.js†L378-L483】【F:src/output/billingOutput.js†L793-L857】【F:src/output/billingOutput.js†L665-L754】
- `resolveAggregateInvoiceDecision_`/`logAggregateDecisionTrace_` による合算判定結果のトレースは、テンプレート生成時にも出力され、ログ監査に利用されるため現行仕様では維持が必要。 【F:src/output/billingOutput.js†L437-L503】【F:src/output/billingOutput.js†L700-L754】

## UI (main.js.html)

### 削除候補
- 現時点で請求UIに未使用の領収・合算表示ロジックは見当たらず（表示用関数はいずれもレンダリング経路から呼び出されている）。

### 要整理・責務不明
- `renderAggregateStatusBadge` と `renderReceiptStatusBadge` が別ステータス（`aggregateStatus` と `receiptStatus`/`billingFinalized`）を同時にバッジ表示する設計のため、同一行に合算待ち／合算予定が重複表示され得る。どちらをユーザー指標にするか整理が必要。 【F:src/main.js.html†L1004-L1077】
- `renderReceiptStatusBadge` は `receiptStatus` が `AGGREGATE` か行が確定済みの場合のみバッジを出すため、`SETTLED` 等他ステータスはUIで不可視。ステータス運用方針と表示条件の意図を擦り合わせたい。 【F:src/main.js.html†L1033-L1070】
- `updateInvoiceModeControls` では PreparedBilling の選択が無いと個別再発行（partial）入力欄を強制的に disable するため、過去のPreparedを読み込む前に再発行IDを入力できない。入力タイミングの制限が仕様か要確認。 【F:src/main.js.html†L238-L260】
- 領収状態の保存 (`persistReceiptStatus`) は確定済み行が1件でもあると `isReceiptEditingLocked` 経由で受付拒否する。合算確定済みと未確定の混在をUI上で編集できない設計が運用要件か確認が必要。 【F:src/main.js.html†L430-L479】【F:src/main.js.html†L187-L214】

### 現行で必須の表示条件
- `renderReceiptControls` で領収状態ドロップダウンは PreparedBilling が選択され、かつローディング中でなく、かつ確定ロックが掛かっていない場合のみ編集可。合算月入力はステータスが `AGGREGATE` の時に限り有効化される。 【F:src/main.js.html†L366-L391】
- `renderFinalizedLockNotice`/`shouldBlockFinalizedBillingOperation` により確定済み行があると再集計・PDF再生成・領収編集がブロックされる。確定フローの安全装置として動作。 【F:src/main.js.html†L187-L214】
- `renderStatusBadges` で行バッジが組み立てられ、請求一覧の氏名列に必ず通過するため、上記表示条件の影響が一覧に反映される。 【F:src/main.js.html†L1033-L1077】【F:src/main.js.html†L216-235】
