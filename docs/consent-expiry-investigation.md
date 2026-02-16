# Consent Expiry Investigation (Issue: resolveConsentExpiry_ 呼び出し最適化)

## 調査結果

### 1. `resolveConsentExpiry_` の呼び出し回数（改修前）
- `getDashboardData` 実行時（管理者表示）
  - `buildDashboardPatients_` 内: **1回/患者**
  - `buildDashboardPatientStatusTags_` 内: **1回/患者**
  - `buildOverviewFromConsent_` 内: **1回/患者**
- 合計: **3回/患者**（表示用途の重複計算あり）

### 2. `getDashboardData` 内での評価回数
- `resolveConsentExpiry_` は患者配列の構築・タグ表示・上段概要の3箇所で評価され、同一患者に対して重複していた。
- `parseConsentDate_`（旧名称）は `resolveConsentExpiry_` の結果を各表示処理で再パースしていた。

### 3. 患者1件あたりの実測（改修後）
- テスト `testConsentExpiryResolutionRunsOncePerPatient` で `resolveConsentExpiry_` をフックし、
  - 患者3件で `resolveConsentExpiry_` 呼び出し数 **3回**
  - すなわち **1回/患者** を確認。

## 改修方針への反映

- 同意期限は患者整形時（`buildDashboardPatients_`）に1回だけ計算。
- 表示ロジック（ステータスタグ、概要表示）では `patient.consentExpiry` を利用し、同意期限の再解決を行わない。
- `parseConsentDate_` は `parseConsentDateInternal_` に改名し、内部利用へ限定。
- 同意期限解析関連のデバッグ/失敗ログを削除。
