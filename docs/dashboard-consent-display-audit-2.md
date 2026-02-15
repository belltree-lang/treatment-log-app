# ダッシュボード② 同意表示不整合 監査レポート

## 監査条件
- 実行時刻 (`now`): `2025-02-01T00:00:00Z`
- 対象ロジック:
  - `buildOverviewFromConsent_`（上段同意ブロック）
  - `buildDashboardPatientStatusTags_`（下段患者タグ）
- 判定キー:
  - `consentExpiry`（または raw の `同意期限` / `同意有効期限`）
  - raw の `同意書取得確認`

## 実患者3名 抽出結果

| patientId | name | consentExpiry 元値 | dashboardParseTimestamp_ 結果 | 同意書取得確認 raw値 | dashboardDaysBetween_ 結果 | 上段表示対象 | 下段consentタグ | 理由 |
|---|---|---|---|---|---:|---|---|---|
| 001 | 山田太郎 | `2025-02-20` | `2025-02-20T00:00:00.000Z` | `""` | `19` | YES | YES | 未取得かつ期限あり。期限日が未来なので上段は「要対応（残19日）」、下段は `consent:要対応`。 |
| 002 | 佐藤花子 | `2025-01-20` | `2025-01-20T00:00:00.000Z` | `""` | `-12` | YES | YES | 未取得かつ期限あり。期限超過のため上段は「期限超過（12日超過）」、下段は `consent:期限超過`。 |
| 003 | 山田花子 | `2025-02-20` | `2025-02-20T00:00:00.000Z` | `済` | `19` | NO | NO | `同意書取得確認` が truthy のため、上段・下段とも同意表示から除外。 |

## `buildOverviewFromConsent_` 全文

```javascript
function buildOverviewFromConsent_(patientInfo, scope, patientNameMap, now) {
  const items = [];
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;
  const targetNow = dashboardCoerceDate_(now) || new Date();

  Object.keys(patientInfo || {}).forEach(pid => {
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    const info = patientInfo[pid] || {};
    const consentExpiry = info.consentExpiry || (info.raw && (info.raw['同意期限'] || info.raw['同意有効期限']));
    const consentExpiryDate = dashboardParseTimestamp_(consentExpiry);
    const consentAcquired = resolvePatientRawValue_(info.raw, ['同意書取得確認']);
    if (consentAcquired || !consentExpiryDate) return;

    const todayStart = new Date(targetNow.getFullYear(), targetNow.getMonth(), targetNow.getDate());
    const expiryStart = new Date(consentExpiryDate.getFullYear(), consentExpiryDate.getMonth(), consentExpiryDate.getDate());
    const diffDays = Math.floor((expiryStart.getTime() - todayStart.getTime()) / (24 * 60 * 60 * 1000));
    const label = diffDays >= 0
      ? `要対応（残${diffDays}日）`
      : `期限超過（${Math.abs(diffDays)}日超過）`;
    const name = info.name || patientNameMap[pid] || '';
    items.push({
      patientId: pid,
      name,
      subText: label
    });
  });

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  return { items };
}
```

## `buildDashboardPatientStatusTags_` 全文

```javascript
function buildDashboardPatientStatusTags_(patient, params, maybeNow) {
  const tags = [];
  const options = params && typeof params === 'object' && !(params instanceof Date)
    ? params
    : { aiReportAt: params, now: maybeNow };
  const targetNow = dashboardCoerceDate_(options.now) || new Date();
  const aiReportAt = options.aiReportAt;
  const consentExpiry = patient && (patient.consentExpiry || (patient.raw && (patient.raw['同意期限'] || patient.raw['同意有効期限'])));
  const consentExpiryDate = dashboardParseTimestamp_(consentExpiry);
  const raw = patient && patient.raw ? patient.raw : null;
  const consentAcquired = resolvePatientRawValue_(raw, ['同意書取得確認']);
  const consentExpired = consentExpiryDate && dashboardDaysBetween_(targetNow, consentExpiryDate, true) <= 0;

  if (!consentAcquired && consentExpiryDate) {
    tags.push({ type: 'consent', label: consentExpired ? '期限超過' : '要対応' });
  }

  const reportDate = dashboardParseTimestamp_(aiReportAt);
  if (!consentAcquired) {
    tags.push({ type: 'report', label: reportDate ? '作成済' : '未作成' });
  }

  return tags;
}
```

## visiblePatientIds.size と包含確認

- `visiblePatientIds.size = 3`
- 含有確認:
  - `001`: 含まれる
  - `002`: 含まれる
  - `003`: 含まれる
- 備考: `004` はログが50日より古く、`visiblePatientIds` から除外。

## 4ケース判定（YES/NO）

| ケース | 上段consentRelatedに出るか | 下段consentタグ出るか |
|---|---|---|
| 1) 期限内未取得 | YES | YES |
| 2) 期限超過未取得 | YES | YES |
| 3) 同意取得確認済 | NO | NO |
| 4) 期限未登録 | NO | NO |

## バグ候補

1. **同日0時の境界判定差**
   - 下段タグは `dashboardDaysBetween_(now, expiry, true) <= 0` を使うため、同日0時を過ぎた時点で `期限超過` になり得る。
   - 上段は `todayStart` / `expiryStart` に丸めて日単位差分を計算するため、同日中は `残0日` 扱い。
   - その結果、同じ患者で上段「要対応（残0日）」・下段「期限超過」が同時に起き得る。

2. **未取得判定が truthy 依存**
   - `同意書取得確認` は空白trim後に非空なら「取得済み」とみなす。
   - `未`, `確認中`, `✕` のような文字列でも truthy なら取得済み扱いになるため、運用入力ゆらぎで誤判定の恐れ。

3. **`dashboardParseTimestamp_` のタイムゾーン依存**
   - `new Date(str)` 依存のため、`YYYY/MM/DD`・`YYYY-MM-DD` 形式の解釈が実行環境タイムゾーン/実装に影響。
   - 厳密な日付比較が必要な同意期限判定でズレ要因になる。

4. **下段 report タグの条件が consent と非対称**
   - 同意取得済みの場合、下段 `report` タグも非表示になる (`if (!consentAcquired)`)。
   - 仕様上「同意取得済みでも報告書作成状態は見せたい」場合は表示欠落。
