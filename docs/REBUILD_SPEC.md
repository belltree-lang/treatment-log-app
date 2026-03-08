# REBUILD_SPEC

## 目的

この文書は `docs/REBUILD_SCOPE.md` で再構築対象とされた 3 領域について、新規実装のための現行仕様を整理したものです。

対象:

- Dashboard
- Billing
- ScheduleCandidateGenerator

前提:

- 既存コードの修正は行わない
- 読み取りと仕様化のみ
- 現行実装の確認結果と、UI が期待している契約をベースに記述する

---

## 1. 機能一覧

## Dashboard

- 施術者ダッシュボード表示
- ログインユーザー別の患者スコープ判定
- 請求未確認サマリ表示
- 同意期限サマリ表示
- 直近訪問タイムライン表示
- 担当患者一覧表示
- 未回収アラート表示
- 患者カード展開時の既読更新

## Billing

- PreparedBilling の対象月一覧取得
- 管理者判定
- 請求データ集計
- 再集計
- 請求編集保存
- 単票 / 一括 PDF 生成
- 銀行引落シート生成
- 銀行引落の未回収履歴反映
- 集計済みデータからの請求計算プレビュー

## ScheduleCandidateGenerator

- 初回受付登録
- 受付一覧表示
- 候補スロット生成
- 仮押さえ / 確定 / 中止 / ステータス更新
- 下書き / ステータス取得
- スタッフ空き状況表示

注記:

- このうち `generateTrialOpportunities` と `waitlistList` は実装を確認できた
- `intakeRegisterLead` など Intake 系の一部 GAS 関数は、UI から参照されているが、このリポジトリ内で定義位置を確認できなかった
- したがって ScheduleCandidateGenerator 仕様は「UI 契約」と「確認済み実装」に分けて扱う

---

## 2. 各機能の目的

## Dashboard の目的

- 施術者または管理者が、担当患者の状態を一覧的に把握する
- 直近対応が必要な患者を優先表示する
- 未回収・同意期限・請求未確認などの業務アラートを一画面に集約する
- ダッシュボードから施術録画面へ遷移する

## Billing の目的

- 月次請求の事前集計結果を固定化する
- 請求額の手修正を保持する
- 請求 PDF を患者別・担当者別に生成保存する
- 銀行引落用シートおよび未回収履歴を管理する
- 請求処理を「集計済みスナップショット」を中心に運用する

## ScheduleCandidateGenerator の目的

- 初回受付情報を収集する
- Waitlist を元に候補日程を提示する
- SlotMaster / Assignment をもとに空き枠候補を算出する
- 受付から仮押さえ・確定までの導線を用意する

---

## 3. UI仕様

## 3.1 画面一覧

| 機能 | 画面 | 役割 |
|---|---|---|
| Dashboard | `src/dashboard.html` | ダッシュボード本体 |
| Billing | `src/billing.html` + `src/main.js.html` | 請求集計・編集・PDF生成・銀行引落 |
| ScheduleCandidateGenerator | `src/intake.html` | 初回受付入力 |
| ScheduleCandidateGenerator | `src/intake_list.html` | 受付一覧と候補生成 |
| ScheduleCandidateGenerator | `src/vacancy.html` | スタッフ空き状況表示 |

## 3.2 画面ごとの操作

### Dashboard

画面:

- `dashboard.html`

主操作:

- 最新情報取得
- 患者一覧の展開
- 申し送り既読化
- 患者画面への遷移

表示ブロック:

- 請求サマリ
- 同意サマリ
- 訪問サマリ
- 未回収アラート一覧
- 当日 / 前日の訪問タイムライン
- 担当患者一覧

必要データ:

- `meta`
  - user
  - generatedAt
  - error
- `overview`
  - invoiceUnconfirmed
  - consentRelated
  - visitSummary
- `todayVisits`
  - today.date
  - today.visits[]
  - previous.date
  - previous.visits[]
- `unpaidAlerts[]`
- `patients[]`
  - patientId
  - name
  - consentExpiry
  - note
  - invoiceUrl
  - statusTags

### Billing

画面:

- `billing.html`
- `main.js.html`

主操作:

- 集計月入力
- 集計実行
- 既存 PreparedBilling の再集計
- 集計済み月選択
- PDF 出力対象選択
  - 保険請求 PDF
  - 自費請求 PDF
- PDF 出力モード選択
  - 一括発行
  - 個別再発行
- 請求データ編集
  - 医療助成
  - オンライン同意
  - 負担割合
  - payerType
  - unitPrice
  - transportAmount
  - carryOverAmount
  - grandTotal
- 変更保存
- PDF生成
- 銀行引落シート生成
- 未回収反映

必要データ:

- preparedMonths[]
- billingAdminInfo
- prepared payload
  - billingMonth
  - preparedAt
  - preparedBy
  - billingJson[]
  - patients
  - bankFlagsByPatient
  - carryOverByPatient
  - billingOverrideFlags
  - files
- bank sheet summary

### ScheduleCandidateGenerator

#### 初回受付画面

画面:

- `intake.html`

主操作:

- 電話 / 訪問モード切替
- Lead 情報入力
  - 名前
  - 電話
  - フリガナ
  - 住所
  - 他サービス利用状況
  - 希望スケジュール（避けたい曜日 / 時間帯）
  - 初回体験希望日時
  - 問診メモ
- URL コピー
- 登録
- 候補検索
- 仮押さえ
- 確定
- 下書き / 状態の読込

必要データ:

- leadId
- mode
- 基本顧客情報
- avoidSlots
- note 群
- 候補一覧

#### 受付一覧画面

画面:

- `intake_list.html`

主操作:

- Waitlist 一覧表示
- 電話/訪問画面への再遷移
- 候補表示
- `generateTrialOpportunities(leadId)`
- ステータス更新
  - scheduling
  - confirmed
  - abandoned

必要データ:

- waitlist rows[]
  - leadId
  - ts
  - name
  - phone
  - address
  - avoidSlotsJson
  - complaint
  - notes
  - status
  - source

#### 空き状況画面

画面:

- `vacancy.html`

主操作:

- スタッフ名クエリを受けて空き状況を表示

必要データ:

- date
- time
- available

---

## 4. データモデル

## 4.1 Dashboard

### 使用スプレッドシート

- 業務メインスプレッドシート
  - `DASHBOARD_SPREADSHEET_ID`

### 使用シート

- `患者情報`
- `施術録`
- `申し送り`
- `AI報告書`
- `未回収履歴`

### 列構造

Dashboard は固定列よりも「ヘッダ名解決」に依存している。

#### 患者情報

主要利用列:

- `患者ID` / `patientId` / `施術録番号`
- `氏名` / `名前` / `患者名`
- `同意年月日`
- その他は `raw` として保持

#### 施術録

主要利用列:

- `タイムスタンプ` / `日時`
- `施術録番号` / `患者ID`
- `氏名`
- `作成者`
- `施術者` / `担当者`
- `メール`
- `担当者ID`
- `所見`
- 自由記述列群

#### 申し送り

実装上の前提列:

- 1列目: TS
- 2列目: 患者ID
- 3列目: ユーザー
- 4列目: メモ
- 5列目: FileIds

#### 未回収履歴

主要利用列:

- `患者ID`
- `氏名`
- `対象月`
- `金額`
- `理由`
- `備考`
- `記録日時`

### 依存関係

- 患者情報が基盤
- 施術録・申し送り・AI報告書・未回収履歴は患者IDに紐づく派生情報
- 請求書 PDF は Drive フォルダから患者名ベースで関連付ける

## 4.2 Billing

### 使用スプレッドシート

- 業務メインスプレッドシート

### 使用シート

明示的に確認できるシート:

- `患者情報`
- `施術録`
- `銀行情報`
- `BillingOverrides`
- `スタッフ一覧`
- `CarryOverLedger`
- `請求履歴`
- `未回収履歴`
- `PreparedBillingMeta`
- `PreparedBillingMetaJson`
- `PreparedBillingJson`
- 月次の `銀行引落_YYYYMM` 系シート

### 列構造

#### PreparedBillingMeta

ヘッダ:

- `billingMonth`
- `preparedAt`
- `preparedBy`
- `payloadVersion`
- `note`

#### PreparedBillingMetaJson

ヘッダ:

- `billingMonth`
- `chunkIndex`
- `payloadChunk`

#### PreparedBillingJson

ヘッダ:

- `billingMonth`
- `patientId`
- `billingRowJson`

#### 請求履歴

ヘッダ:

- `billingMonth`
- `patientId`
- `nameKanji`
- `billingAmount`
- `carryOverAmount`
- `grandTotal`
- `paidAmount`
- `unpaidAmount`
- `bankStatus`
- `updatedAt`
- `memo`
- `receiptStatus`
- `aggregateUntilMonth`
- `previousReceiptAmount`

#### 未回収履歴

ヘッダ:

- `patientId`
- `対象月`
- `金額`
- `理由`
- `備考`
- `記録日時`

#### BillingOverrides

コード上の主要列:

- `billingMonth`
- `patientId`
- `entryType`
- `manualUnitPrice`
- `manualTransportAmount`
- `carryOverAmount`
- `adjustedVisitCount`
- `manualSelfPayAmount`
- `manualBillingAmount`

#### 患者情報

請求で重要な列:

- 患者ID
- 氏名
- フリガナ
- 負担割合
- 同意情報
- 自費 / 保険関連属性
- 銀行口座情報

### 依存関係

- `患者情報` が基本マスタ
- `施術録` が visit count ソース
- `BillingOverrides` が手修正ソース
- `銀行情報` が引落出力ソース
- `CarryOverLedger` / `未回収履歴` / `請求履歴` が繰越・回収履歴のソース
- `PreparedBilling*` が月次の固定スナップショット

## 4.3 ScheduleCandidateGenerator

### 使用スプレッドシート

- スケジュール専用スプレッドシート
  - `SCHEDULE_SS_ID`

### 使用シート

- `Waitlist`
- `SlotMaster`
- `Assignment`
- `TrialOpportunities`
- `Config`
- `GeoCache`

### 列構造

#### Waitlist

`waitlistList()` の返却形から推定できる主要列:

- `leadId`
- `ts`
- `name`
- `phone`
- `address`
- `avoidSlotsJson`
- `complaint`
- `notes`
- `status`
- `source`

#### SlotMaster

`setupTestSlotMaster()` と `generateTrialOpportunities()` から確認できる列:

- `slot_id`
- `staff`
- `weekday`
- `start_time`
- `duration_min`
- `visit_type`
- `area`
- `active`
- `priority`

#### Assignment

`generateTrialOpportunities()` で参照される主要列:

- `slot_id`
- `effective_from`
- `effective_to`
- `status`

#### TrialOpportunities

`generateTrialOpportunities()` が書き込む項目:

- `date`
- `staff`
- `area`
- `gap_start`
- `gap_end`
- `gap_minutes`
- `candidate_patient_id`
- `candidate_name`
- `candidate_address`
- `travel_prev_min`
- `service_min`
- `travel_next_min`
- `total_required_min`
- `slack_min`
- `reason`
- `note`

#### Config

確認できるキー:

- `look_ahead_days`
- `service_duration_min`
- `timezone`
- `force_zero_travel`

#### GeoCache

列:

- `address`
- `lat`
- `lng`

### 依存関係

- `Waitlist` が候補生成の入力
- `SlotMaster` が空き枠マスタ
- `Assignment` が既存アサイン状態
- `Config` が候補生成パラメータ
- `GeoCache` が地理情報キャッシュ
- `TrialOpportunities` が候補出力先

---

## 5. API仕様

## 5.1 Dashboard API

### `getDashboardData(options)`

役割:

- ダッシュボード描画に必要な全データを返す統合 API

入力:

- `options.user` 任意
- `options.mock` 任意
- `options.now` 任意
- 内部注入用の `patientInfo`, `notes`, `aiReports`, `invoices`, `treatmentLogs` 等

出力:

```json
{
  "tasks": [],
  "todayVisits": {
    "today": { "date": "YYYY-MM-DD", "visits": [] },
    "previous": { "date": "YYYY-MM-DD|null", "visits": [] }
  },
  "patients": [],
  "unpaidAlerts": [],
  "warnings": [],
  "overview": {},
  "meta": {
    "generatedAt": "ISO8601",
    "user": "email",
    "setupIncomplete": false,
    "error": ""
  }
}
```

### `markAsRead({ patientId, email?, readAt? })`

役割:

- 患者カード展開時に申し送り既読時刻を保存

入力:

- `patientId`
- `email` 任意
- `readAt` 任意

出力:

```json
{
  "ok": true,
  "patientId": "PID",
  "readAt": "ISO8601"
}
```

## 5.2 Billing API

### `getPreparedBillingMonths()`

役割:

- 集計済み月一覧を返す

入力:

- なし

出力:

- `["202602","202601", ...]`

### `getBillingAdminInfo()`

役割:

- ログインユーザーの請求管理者判定

出力:

```json
{
  "isAdmin": true,
  "email": "user@example.com"
}
```

### `prepareBillingData(billingMonth)`

役割:

- 請求データを集計し PreparedBilling 相当の payload を返す

入力:

- `billingMonth`: `YYYY-MM` または `YYYYMM`

出力:

- PreparedBilling payload

### `resetPreparedBillingAndPrepare(billingMonth)`

役割:

- 指定月の既存 PreparedBilling を削除して再集計

入力:

- `billingMonth`

出力:

- PreparedBilling payload

### `applyBillingEdits(billingMonth, options)`

役割:

- 患者情報編集と BillingOverrides 編集を保存し、再集計 payload を返す

入力:

```json
{
  "patientInfoUpdates": [
    { "patientId": "PID", "medicalAssistance": 1, "onlineConsent": 0, "burdenRate": 1, "payerType": "保険" }
  ],
  "billingOverridesUpdates": [
    { "patientId": "PID", "manualUnitPrice": 417, "manualTransportAmount": 33, "carryOverAmount": 0, "adjustedVisitCount": 5, "manualSelfPayAmount": "", "manualBillingAmount": "" }
  ]
}
```

出力:

- PreparedBilling payload

### `calculateBillingRowTotalsServer(row)`

役割:

- UI プレビュー用の請求行再計算

入力:

- 単一請求行オブジェクト

出力:

```json
{
  "visitCount": 0,
  "treatmentUnitPrice": 0,
  "treatmentAmount": 0,
  "transportAmount": 0,
  "carryOverAmount": 0,
  "billingAmount": 0,
  "manualSelfPayAmount": 0,
  "grandTotal": 0
}
```

### `generatePreparedInvoicesForMonth(billingMonth, options)`

役割:

- 集計済み payload をもとに PDF を生成し Drive へ保存

入力:

```json
{
  "patientInfoUpdates": [],
  "billingOverridesUpdates": [],
  "invoiceMode": "bulk|partial",
  "invoicePatientIds": ["PID1", "PID2"],
  "includeInsurancePdf": true,
  "includeSelfPayPdf": false
}
```

出力:

- PreparedBilling payload + `files` 等の出力情報

### `generateSimpleBankSheet(billingMonth)`

役割:

- `銀行情報` をコピーして簡易銀行引落シートを作る

出力:

```json
{
  "billingMonth": "YYYYMM",
  "sheetName": "銀行引落_YYYYMM",
  "rows": 0,
  "filled": 0,
  "missingAccounts": []
}
```

### `generateBankWithdrawalSheetFromCache(billingMonth)`

役割:

- 集計済み PreparedBilling を元に本番用銀行引落シートを同期

出力:

- `billingMonth`
- `billingCount`
- `preparedAt`
- `sheetSummary`

### `applyBankWithdrawalUnpaidFromUi(billingMonth)`

役割:

- 銀行引落シート上の未回収フラグを `未回収履歴` へ反映

出力:

- `billingMonth`
- `sheetSummary`
- `checkedRows`
- `added`
- `skipped`

## 5.3 ScheduleCandidateGenerator API

## 確認済み実装

### `waitlistList()`

役割:

- Waitlist 一覧取得

出力:

```json
[
  {
    "leadId": "",
    "ts": "",
    "name": "",
    "phone": "",
    "address": "",
    "avoidSlotsJson": "",
    "complaint": "",
    "notes": "",
    "status": "open",
    "source": ""
  }
]
```

### `generateTrialOpportunities(leadId)`

役割:

- Waitlist / SlotMaster / Assignment / Config を使って候補生成
- `TrialOpportunities` に書き込み

入力:

- `leadId`

出力:

```json
[
  {
    "staff": "",
    "area": "",
    "start": "YYYY-MM-DD HH:mm",
    "end": "YYYY-MM-DD HH:mm",
    "leadId": ""
  }
]
```

## UI 契約として確認できるが、定義位置未確認

以下は HTML から呼ばれているが、このリポジトリ内で定義位置を確認できなかった関数:

- `intakeRegisterLead(payload)`
- `intakeSuggestSlots(lead, opts)`
- `intakeHoldSlot(lead, start, end, staffEmail)`
- `intakeConfirm(lead)`
- `intakeLoadDraft(lead)`
- `intakeGetStatus(lead)`
- `intakeAbandon(lead, reason)`
- `intakeUpdateStatus(lead, status)`
- `vacancyList(staff)`

新規実装時は、UI が既に期待しているため、この名称・引数・基本返却構造を維持するのが妥当です。

---

## 6. 処理フロー

## 6.1 Dashboard

```text
UI (dashboard.html)
  -> getDashboardData(options)
  -> Spreadsheet:
       患者情報
       施術録
       申し送り
       AI報告書
       未回収履歴
     Drive:
       請求書フォルダ
  -> 集約済み JSON
  -> UI 描画

患者カード展開
  -> markAsRead({ patientId })
  -> Script Properties (HANDOVER_LAST_READ)
  -> UI の unread 状態更新
```

## 6.2 Billing

```text
UI (billing.html / main.js.html)
  -> prepareBillingData(month)
  -> Spreadsheet:
       患者情報
       施術録
       銀行情報
       BillingOverrides
       CarryOverLedger
       未回収履歴
       請求履歴
  -> PreparedBilling payload 生成
  -> PreparedBillingMeta / MetaJson / Json 保存
  -> UI テーブル表示

UI 編集
  -> applyBillingEdits(month, payload)
  -> 患者情報 / BillingOverrides 更新
  -> PreparedBilling 再生成・再保存
  -> UI 再描画

PDF 生成
  -> generatePreparedInvoicesForMonth(month, options)
  -> PreparedBilling 読込
  -> invoice_template.html 等で PDF 作成
  -> Drive 保存
  -> 必要に応じて 請求履歴 / 付随出力更新
  -> UI 結果表示

銀行引落
  -> generateBankWithdrawalSheetFromCache(month)
  -> 銀行引落_YYYYMM シート生成 / 同期
  -> applyBankWithdrawalUnpaidFromUi(month)
  -> 未回収履歴 更新
  -> UI 反映
```

## 6.3 ScheduleCandidateGenerator

```text
UI (intake.html)
  -> intakeRegisterLead(payload)
  -> Waitlist 登録
  -> leadId を URL に反映

UI (intake_list.html)
  -> waitlistList()
  -> Waitlist 一覧表示
  -> generateTrialOpportunities(leadId)
  -> Waitlist + SlotMaster + Assignment + Config + GeoCache 参照
  -> TrialOpportunities へ候補書込
  -> UI 候補表示

候補確定系
  -> intakeHoldSlot / intakeConfirm / intakeUpdateStatus / intakeAbandon
  -> Waitlist / Assignment 等を更新
  -> UI 再読み込み

空き状況
  -> vacancyList(staff)
  -> スタッフ別空き状況取得
  -> UI 表示
```

---

## 7. 副作用

## 7.1 Drive

### Dashboard

- 請求書 PDF フォルダの列挙
- ファイル URL 読み取り

### Billing

- 請求書 PDF の保存
- 担当者別フォルダへの保存
- 入金結果 PDF の読込

### ScheduleCandidateGenerator

- 現時点で Drive 副作用は確認できない

## 7.2 PDF

### Billing

- 請求書 PDF 生成
- 銀行引落関連出力の派生
- 入金結果 PDF の解析

### Dashboard

- PDF 自体は生成しない
- Drive 上の PDF を参照する

### ScheduleCandidateGenerator

- PDF 副作用は確認できない

## 7.3 外部API

### Dashboard

- 外部 API 呼び出しは直接確認できない

### Billing

- Billing 自体の主要フローでは外部 API 呼び出しは確認していない
- ただし同一リポジトリ内には OpenAI / webhook 連携が存在する

### ScheduleCandidateGenerator

- Google Maps Geocoder
  - 住所 -> 緯度経度変換
- GeoCache へキャッシュ

---

## 8. 非スコープ

この仕様書では以下を対象外とする。

- 施術録本体
  - `src/app.html`
  - `src/Code.js` の施術録保存 / 更新 / 削除 / News / 申し送り本体
- 勤怠
  - `src/attendance.html`
  - `VisitAttendance` 系ロジック
- 給与
  - `src/payroll.html`
  - Payroll 系ロジック
- Albyte
  - `src/albyte.html`
  - `src/albyte_admin.html`
  - `src/albyte_report.html`
  - Albyte 系ロジック

---

## 付記

本仕様書は「現行観測ベース」であり、以下の 2 種類の情報を含みます。

1. 実装を確認できた事実
2. UI が期待している契約だが、定義位置未確認のもの

新規実装時は、この差を前提に、

- 既存 UI 契約を維持する箇所
- 現行サーバ実装から移植すべき箇所

を分離して設計する必要があります。

