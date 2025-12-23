# 請求「確定」状態の表現方針（Issue #829 設計メモ）

## 目的
- 銀行引落結果（未回収/合算）や `receiptStatus` を併用している現行ロジックに、「請求が確定した」状態を矛盾なく追加する。
- 既存データ互換性を保ったまま、確定後の再計算・上書き可否を明確にするメタデータを定義する。
- 本ドキュメントは設計案のみで、実装やテストの変更は行わない。

## 現行フィールドと課題
- `receiptStatus`: `''` / `UNPAID` / `HOLD` / `AGGREGATE` / `PAID` などを許容し、領収書表示や備考生成に利用される。合算指定の有無で `showReceipt` が変化する。【F:docs/billing-receipt-status.md†L1-L140】
- `aggregateUntilMonth`: 合算終了月。`receiptStatus` が `AGGREGATE` のときにのみ正規化・保持され、それ以外は空文字にリセットされる。【F:docs/billing-receipt-status.md†L142-L172】
- 銀行引落シートの `AE`（未回収）/`AF`（合算）フラグ: 今後はこれを真とみなし、`receiptStatus` / `aggregateUntilMonth` を派生値として上書きする方針。【F:docs/bank-withdrawal-flags-design.md†L1-L65】
- 課題: 「請求が確定して以降は再計算・上書きしない」境界が明確でなく、`appendBillingHistoryRows` も既存値優先であるため、手動編集や履歴残存と競合しやすい。

## 表現モデルの選択肢
1. **`receiptStatus` を拡張して `FINALIZED` を追加する**
   - 長所: 既存フィールドに寄せられるため UI/保存先を増やさずに済む。
   - 短所: 領収書表示ロジックと混在し、`FINALIZED + AF` のような組み合わせの意味が不明瞭になる。過去データでは空/UNPAID が確定かどうか判別できない。
2. **確定専用のメタフィールドを追加する（推奨）**
   - 例: `billingFinalized: boolean`, `finalizedAt: string (ISO)`, `finalizedBy: string`, `finalizationSource: 'manual' | 'auto'`。
   - 長所: 領収可否ロジックと分離し、AE/AF や `receiptStatus` の再計算と衝突しない。確定後の上書き制御や監査ログを素直に持てる。
   - 短所: 新フィールドの保存先（Prepared/History シート、meta JSON）を追加する必要がある。

## 推奨モデルと流通先
- `billingFinalized: boolean`（必須）: 請求月 × 患者 ID 単位で「確定済みか」を保持。デフォルト `false`。
- `finalizedAt: string`（任意）: ISO 形式のタイムスタンプ。Apps Script 実行時刻を保存。
- `finalizedBy: string`（任意）: GAS 実行ユーザー（メールアドレス）。
- `finalizationSource: 'manual' | 'auto'`（任意）: UI 操作か自動処理（例: 銀行結果取り込み後の一括確定）かを記録。
- 保存先: `PreparedBillingJson` / `PreparedBillingMetaJson` / 履歴シートに新カラムを追加し、月×患者行に同値を持たせる。既存行は空扱いで読み込み時に `false` へフォールバックする。

## 互換性と移行方針
- **既存データの解釈**: `billingFinalized` が無い場合は `false` とみなし、従来通り再計算・上書き可能にする。過去の `receiptStatus` / `aggregateUntilMonth` はそのまま温存。
- **AE/AF との整合**: 確定済みでも `bankFlags` の読込自体は許容するが、`billingFinalized` が `true` の行は `receiptStatus` / `aggregateUntilMonth` へ派生書き込みを行わない（表示用には読み出しのみ）。
- **UI/操作**: 「請求を確定する」チェックボックス（またはボタン）を追加し、確定時に上記メタデータを設定。確定解除が必要な場合は明示的に `billingFinalized` を `false` に戻す UI を用意し、解除時のみ AE/AF → receiptStatus 再計算を許可する。
- **履歴への影響**: `appendBillingHistoryRows` は `billingFinalized` が `true` の行を上書きしないガードを入れることで、手動確定後の自動再計算を防止。未確定行は従来通り最新値を反映する。

## 処理フロー例（推奨案）
1. prepared 生成時に銀行引落シートから AE/AF を読み込み、`bankFlags` を月×患者キーで保持する（既存方針）。【F:docs/bank-withdrawal-flags-design.md†L18-L65】
2. AE/AF から `receiptStatus` / `aggregateUntilMonth` を算出し、`billingFinalized !== true` の行に限って上書きする。既存の手入力値は確定済みフラグが無い限り上書きされる点を明文化する。
3. UI で「確定」操作が行われたら、対象行の `billingFinalized` を `true` にし、`finalizedAt` / `finalizedBy` / `finalizationSource` を保存する（Prepared/履歴/メタ JSON へ展開）。
4. 確定済み行を再表示する際は `billingFinalized` を表示ロックの根拠としつつ、`receiptStatus` / `aggregateUntilMonth` は読み取り専用にする。解除後にのみ再編集・再計算を許可。

## 追加で決めるべき事項
- 確定解除の権限・監査要件（誰がいつ解除したかをどこに残すか）。
- スプレッドシート列追加の命名（例: `Finalized?`, `FinalizedAt`, `FinalizedBy`）。
- `billingFinalized` を true にした時点の `receiptStatus` を「凍結値」として別列に写すか（履歴比較用）。
- バッチ系（請求一括生成、PDF 出力）で「未確定のみ処理する」フィルタを入れるかどうか。

本メモは Issue #829 の前提となる設計案として共有する。実装時は既存の `receiptStatus` / `aggregateUntilMonth` / `bankFlags` との整合を優先し、未確定データへの後方互換を保ったままメタデータを追加する。
