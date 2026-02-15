# Issue #1608 調査メモ（原因切り分けのみ）

## 結論（A〜E）

- **A. visiblePatientIds（担当スコープ）で除外**: **YES**
  - staff ユーザーでは `visiblePatientIds` が「担当者一致した施術ログのうち直近50日」に限定されるため、同意期限が近い患者でも条件を満たさないと上段/下段とも表示対象外になる。
- **B. consentExpiryが空/列解決失敗/値形式でparseできず除外**: **YES**
  - 同意期限列の候補に合わないヘッダー、または `new Date(str)` で解釈不能な文字列の場合、同意期限扱いにならず上段/下段とも除外される。
- **C. raw['同意書取得確認'] が truthy 扱いで誤って取得済み判定**: **YES（最有力）**
  - `resolvePatientRawValue_` は「空でない文字列」を返すため、`"未取得"`, `"FALSE"`, `"0"` でも truthy となり `consentAcquired=true` 扱いになって上段非表示・下段タグ非表示が起きる。
- **D. 日付境界/TZ/丸め差でラベル判定が想定とズレ**: **YES（表示有無よりラベル不整合寄り）**
  - 上段は日付開始時刻同士で `diffDays` を算出、下段は時刻込み差分を `Math.floor` で算出しており、同日境界で「残0日/期限超過」の判定がズレ得る。
- **E. 上段と下段で参照キーや前提がズレている**: **NO（大きなキー不一致は未確認）**
  - 両者とも `consentExpiry` / `raw['同意期限'|'同意有効期限']` と `raw['同意書取得確認']` を同様に参照しており、根本的な参照キー不一致は見当たらない。

## 判定経路（コード上）

1. staff ユーザーでは、担当一致ログから `responsiblePatientIds` を作り `visiblePatientIds` として利用。
2. `buildDashboardOverview_` は `allowedPatientIds` を `buildOverviewFromConsent_` に渡し、上段②同意をフィルタ。
3. `buildDashboardPatients_` も同じ `allowedPatientIds` で患者一覧をフィルタし、下段タグ `buildDashboardPatientStatusTags_` を作成。
4. 上段/下段とも `consentExpiry` の parse 失敗時は対象外。
5. 上段/下段とも `resolvePatientRawValue_(raw, ['同意書取得確認'])` が truthy なら「取得済み」とみなして同意表示を作らない。

## 最小限の一時ログ案（原因特定用）

> すべて `dashboardLogContext_('consent-debug', JSON.stringify({...}))` 形式で1行JSON。調査後に削除。

- 追加位置1: `getDashboardData` で `visiblePatientIds` 決定直後
  - 出力例:
  - `{ "phase":"scope", "user":"...", "isAdmin":false, "matchedLogs":123, "visiblePatientIds":45, "sample":["P001","P002"] }`
- 追加位置2: `buildOverviewFromConsent_` の各患者ループ内（`continue`直前のみ）
  - 出力例:
  - `{ "phase":"overview", "pid":"P001", "inScope":true, "consentExpiryRaw":"2025/10/01", "consentExpiryParsed":"2025-10-01T00:00:00.000Z", "consentAcquiredRaw":"FALSE", "consentAcquiredTruthy":true, "skipReason":"consentAcquired" }`
- 追加位置3: `buildDashboardPatientStatusTags_` の return 直前
  - 出力例:
  - `{ "phase":"tag", "pid":"P001", "consentExpiryRaw":"2025/10/01", "consentExpiryParsed":"2025-10-01T00:00:00.000Z", "consentAcquiredRaw":"FALSE", "consentAcquiredTruthy":true, "consentExpired":false, "tagTypes":["report"] }`

## 患者ID単位チェックリスト

対象患者ごとに以下を確認する。

1. `visiblePatientIds` に含まれるか（staff時のみ）
2. `consentExpiry` 元値
   - `patient.consentExpiry`
   - どのヘッダー列から来たか（同意期限/同意書期限/同意有効期限/同意期限日）
3. `dashboardParseTimestamp_(consentExpiry)` の結果（Date/null）
4. `raw['同意書取得確認']` の元値と判定
   - 元値（例: `true`, `false`, `"FALSE"`, `"未取得"`, 空）
   - 現在実装での truthy/falsy 判定
5. 上段②同意
   - 表示/非表示
   - ラベル（要対応（残N日）/期限超過（N日超過））
6. 下段患者タグ
   - `consent` タグ有無
   - ラベル（要対応/期限超過）

## 最小修正方針（1本に限定）

**方針: `同意書取得確認` を厳密に正規化して boolean 化する。**

- 具体:
  - `resolvePatientRawValue_` の戻り値をそのまま truthy 判定せず、同意取得判定専用ヘルパー（例: `isConsentAcquired_`）で
    - `true`, `"true"`, `"済"`, `"取得済"`, チェックONのみを `true`
    - `false`, `"false"`, `"未"`, `"未取得"`, 空を `false`
  - 上段 `buildOverviewFromConsent_` と下段 `buildDashboardPatientStatusTags_` の両方で同一ヘルパーを使用。

- 期待効果:
  - 「上段に出るべき患者が出ない」「下段に同意タグが出ない」を同時に起こす主要因（文字列truthy誤判定）を最小変更で抑止できる。
