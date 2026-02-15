# Consent Expiry Data-Layer Investigation

## 1) resolveConsentExpiry_ 実装確認
- 参照キー（優先順）
  1. `info.consentExpiry`
  2. `raw['同意期限']`
  3. `raw['同意有効期限']`
  4. `raw['同意期限日']`
- `raw['consentExpiry']` は参照していない。
- `patientInfo` の `patients[pid]` オブジェクトを受け取り、`patient.raw` を読む。
- `null` 扱い条件:
  - 値が `null/undefined`
  - 文字列かつ trim 後に空文字

## 2) loadPatientInfo のカラムマッピング確認
- ヘッダは `getDisplayValues()` で取得。
- 同意期限列の候補: `['同意期限', '同意書期限', '同意有効期限', '同意期限日']`。
- マッチ方式:
  - ヘッダは trim + lowercase で正規化
  - 候補側も trim + lowercase で正規化
  - 完全一致比較
- index は 1-origin（見つからない場合は 0）。
- `raw` オブジェクト生成:
  - 全ヘッダを走査し、`raw[String(header).trim()] = row[idx]` を格納。

## 3) 実シート構造確認
- ローカルリポジトリ環境では Google Apps Script 実行コンテキスト・実スプレッドシートへの接続がないため未実施。

## 4) PID単体検証
- 実シートデータにアクセスできないため未実施。

## 5) parseConsentDate_ の挙動
- 受理フォーマット:
  - `Date` オブジェクト（有効日付）
  - `yyyy-M-d` / `yyyy-MM-dd`
  - `yyyy/M/d` / `yyyy/MM/dd`
  - `yyyy年M月d日`
  - ISO日時（例: `2025-02-01T00:00:00Z`）
- `null` になる主な条件:
  - `null/undefined`
  - 空文字
  - 上記いずれのフォーマットにも一致しない
  - 数値として不正な日付（例: 2025-02-30）
