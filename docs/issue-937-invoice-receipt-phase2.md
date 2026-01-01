# Issue 937 請求書／領収書ロジック Phase2 仕様判断（調査用）

## 背景
- 請求書／領収書ロジックの Phase2 で実装検討中の論点について、現行コードの挙動と仕様としての妥当性を Yes/No で整理する調査 Issue の下書き。
- 本メモでは挙動変更は行わず、開発チケット化に必要な判断メモのみをまとめる。

## 判断サマリ
- **hasPreviousReceiptSheet のデフォルト挙動: No** — 値未設定時に `true` を返すため、実際に前月領収シートがないケースまで表示許可になる。未設定を「未収集」と区別したい場合に破綻するため仕様としては NG と判断。【F:src/output/billingOutput.js†L360-L370】
- **previousReceiptAmount による合算判定: No** — `previousReceiptAmount` が 0 超なら合算対象月が 1 件でも `isAggregateInvoice=true` となるため、金額入力だけで合算モードに入る。合算月の明示がない状態での自動合算は仕様意図と齟齬があると判断。【F:src/output/billingOutput.js†L437-L483】
- **fallbackReceiptMonths / trace の要否: No** — `fallbackReceiptMonths` は決定ロジックに使われず、trace でも常に `fallbackReceiptMonthsUsedInDecision=false` のまま出力される。意思決定に寄与しないデータは仕様として不要、もしくは決定ロジックへ組み込む要検討。【F:src/output/billingOutput.js†L437-L503】
- **aggregateStatus と receiptStatus の役割分離: No** — 一覧 UI で両ステータスのバッジが並立し、同じ合算状態を二重表現するケースがある。どちらをユーザー指標とするか明確な役割分離が必要で、現状は仕様未確定と判断。【F:src/main.js.html†L1004-L1100】

## 参考観察
- `hasPreviousReceiptSheet` のデフォルト `true` は領収シート有無が未入力でも表示許可になるため、データ欠落を検知できない。【F:src/output/billingOutput.js†L360-L379】
- 合算判定では `previousReceiptAmount` が decisionSources に追加され `isAggregateInvoice` を強制するため、領収金額が入っているだけで合算 PDF が生成されうる。【F:src/output/billingOutput.js†L437-L503】
- `fallbackReceiptMonths` は receipt 決定から引き継ぐが、実際の `aggregateDecisionMonths` には使われないまま trace に残るのみ。【F:src/output/billingOutput.js†L460-L483】
- UI 側では `aggregateStatus`（未回収/合算シグナル）と `receiptStatus`（合算予定/確定）双方をバッジ表示し、同一行に「合算待ち」と「合算予定」などが同時に出る状態。【F:src/main.js.html†L1004-L1100】

## 次のアクション案
- 上記 Yes/No 判断を Issue #937 として起票し、必要に応じて仕様修正タスク（デフォルト値の見直し、合算判定条件の整理、trace 廃止または活用、UI ステータス責務の明文化）を分割する。
