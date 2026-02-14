# ダッシュボード・スコープ整合性 調査メモ

## 調査対象コード
- `getDashboardData`（`src/dashboard/api/getDashboardData.js`）
- `buildDashboardPatients_` / `buildDashboardPatientStatusTags_`（同ファイル）
- `buildDashboardOverview_` / `buildOverviewFromInvoiceUnconfirmed_`（同ファイル）
- `assignResponsibleStaff`（`src/dashboard/data/assignResponsibleStaff.js`）

## 事実確認（コードベース）

### 1) visiblePatientIds
- `isAdmin = Boolean(user && user.role === 'admin')`
- `staffMatchedLogs` から **50日以内** の `patientId` のみ `responsiblePatientIds` に追加。
- `visiblePatientIds = isAdmin ? null : responsiblePatientIds`。

### 2) buildDashboardPatients_ のスコープ
- `buildDashboardPatients_(..., allowedPatientIds)` は `basePatients` の全キーを走査。
- ただし `addPatient` 冒頭で `if (allowedPatientIds && !allowedPatientIds.has(patientId)) return;` を実行するため、
  `allowedPatientIds` がある場合はそこで除外される。
- `statusTags` は `addPatient` の内部で作成されるため、除外後の患者に対してのみ計算される。

### 3) statusTags（reportタグ含む）の母集団
- `buildDashboardPatientStatusTags_` は `buildDashboardPatients_` の `addPatient` 内から 1患者につき1回呼ばれる。
- reportタグ（`type: 'report'`）は `aiReportAt` が未設定、または 180日以上前で付与される。
- 従って、staffユーザー時は `allowedPatientIds` で通過した患者だけが判定対象になる。

### 4) overview.invoiceUnconfirmed の母集団
- `buildDashboardOverview_` に `allowedPatientIds: visiblePatientIds` を渡している。
- `buildOverviewFromInvoiceUnconfirmed_` では `allowedPatientIds` があるとき、
  前月患者抽出で `if (allowedPatientIds && !allowedPatientIds.has(pid)) { filteredCount += 1; return; }` を実行。
- その後の `prevMonthPatientIds` を起点に請求未確認候補を作るため、ここでも staffユーザー時は可視患者のみが母集団となる。

### 5) assignResponsible の挙動
- `assignResponsibleStaff` は `patients` の全IDに対して `responsible[pid]` を作る（値は `lastStaffByPatient[pid]` または `null`）。
- さらに `lastStaffByPatient` にのみ存在するIDも追加する。
- ここで作るのは **表示スコープではなく担当者辞書** であり、`visiblePatientIds` の生成には使われていない。

## YES/NO 判定
- statusTags の算出は visiblePatientIds に限定されているか？ **YES**（staff時）
- overview.invoiceUnconfirmed は visiblePatientIds を使っているか？ **YES**（staff時）
- reportタグ算出に全患者を参照していないか？ **YES**（staff時は限定、admin時は全患者）
- assignResponsible が全患者にtrueを立てていないか？ **YES**（true/falseを立てる実装ではなく、`staff|null` を患者ID辞書として作成）

## 補足（数値の取得可否）
- 本リポジトリの静的コード調査では、実運用データの `visiblePatientIds.size` / `reportTagged` 実ID は確定できない。
- ただし以下ログは既に実装済み：
  - `getDashboardData:staffMatchedLogs`
  - `getDashboardData:matchedPatientIds`
  - `getDashboardData:assignResponsible`
  - `billing-debug [billing-scope] visiblePatientIds size=... filteredCount=...`
  - `billing-debug prevMonthPatientIds count=...`
  - `billing-debug pendingPatients count=... sample=...`
