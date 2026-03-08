# SYSTEM_MAP

## 目的

この文書は `treatment-log-app` の現状構造を読み取りベースで可視化し、再構築対象を限定するための分析資料として作成したものです。

- 対象: リポジトリ全体
- 方針: 読み取りと文書化のみ
- 非対象: リファクタリング、命名変更、共通化、保存方式変更、既存ロジック修正

## ディレクトリツリー

```text
treatment-log-app/
├─ .github/
│  └─ workflows/
│     └─ deploy.yml
├─ docs/
│  ├─ 既存の設計メモ・監査メモ・issue調査メモ
│  ├─ SYSTEM_MAP.md
│  └─ REBUILD_SCOPE.md
├─ src/
│  ├─ Code.js
│  ├─ entrypoint.gs
│  ├─ main.gs
│  ├─ utility.js
│  ├─ schedule.js
│  ├─ debug_expose.js
│  ├─ *.html
│  ├─ get/
│  │  └─ billingGet.js
│  ├─ logic/
│  │  ├─ billingLogger.js
│  │  └─ billingLogic.js
│  ├─ output/
│  │  └─ billingOutput.js
│  └─ dashboard/
│     ├─ config.gs
│     ├─ main.gs
│     ├─ api/
│     ├─ auth/
│     ├─ data/
│     ├─ tests/
│     └─ utils/
├─ tests/
│  └─ billing / dashboard / treatment / attendance 関連テスト
├─ .clasp.json
├─ AGENTS.md
└─ README.md
```

## 全体像

このリポジトリには、大きく 4 つの系統が同居しています。

1. `src/Code.js` を中心とした旧来の一体型 GAS アプリ
2. `src/main.gs` と `src/get|logic|output` を中心とした請求処理の分離系
3. `src/dashboard/*` を中心としたダッシュボード分離系
4. `src/schedule.js` を中心とした別スプレッドシート依存のスケジュール候補生成系

現状は「完全分離済み」ではなく、「一部を分離し始めたが、依然としてグローバルスコープと既存モノリスに強く依存している」状態です。

## GAS エントリーポイント

### 1. `src/Code.js`

- `doGet(e)`
  - `view` パラメータで HTML を切り替える旧来のメインルータ
  - ルーティング先:
    - `welcome`
    - `intake`
    - `visit`
    - `intake_list`
    - `admin`
    - `attendance`
    - `vacancy`
    - `albyte`
    - `albyte_admin`
    - `albyte_report`
    - `payroll`
    - `payroll_pdf_family`
    - `billing`
    - `dashboard`
    - `record` -> `app.html`
    - `report`
- `shouldHandleDashboardApi_(e)`
  - ダッシュボード API を JSON レスポンスとして返す分岐を内包

### 2. `src/entrypoint.gs`

- `doGet(e)`
  - `handleDashboardDoGet_`
  - `handleBillingDoGet_`
  - 上記の存在に依存して委譲する最小ルータ

### 3. `src/main.gs`

- `doGet(e)`
  - `handleBillingDoGet_(e)` に直接委譲
- `handleBillingDoGet_(e)`
  - `billing.html` を返す請求アプリ入口
- `onOpen()`
  - スプレッドシートメニューから請求処理・勤怠同期を起動

### 4. `src/dashboard/main.gs`

- `handleDashboardDoGet_(e)`
  - ダッシュボード HTML または JSON API を返す
- `shouldHandleDashboardRequest_(e)`
- `shouldHandleDashboardApi_(e)`

## エントリーポイントの所見

- `doGet` が 3 箇所に存在する
  - `src/Code.js`
  - `src/entrypoint.gs`
  - `src/main.gs`
- Apps Script はトップレベル関数がグローバル共有されるため、デプロイ単位によっては競合余地がある
- どの `doGet` を正とするかが、リポジトリだけでは一意に見えない
- 「再構築前にまず境界を固定すべき箇所」はここ

## HTML ごとの利用目的

### ルータ起点の画面

| HTML | 主用途 | 主な接続先 |
|---|---|---|
| `welcome.html` | メニュー画面 | `?view=...` 遷移 |
| `app.html` | 施術録入力、News、施術記録一覧、報告書作成ヒント | `src/Code.js` の施術録/申し送り/AI報告書関連 |
| `attendance.html` | 訪問スタッフ勤怠、修正申請、有給申請、管理者承認 | `getVisitAttendancePortalData`, `submitVisitAttendanceRequest`, `previewPaidLeavePlan`, `submitPaidLeaveRequest`, `updateVisitAttendanceRequestStatus` |
| `billing.html` | 請求集計、PreparedBilling 選択、PDF 生成 | `src/main.gs` + `src/get|logic|output` |
| `dashboard.html` | 施術者ダッシュボード | `getDashboardData`, `markAsRead` |
| `payroll.html` | 給与マスタ、給与明細一括生成、保険料設定 | `src/Code.js` の給与管理系 |
| `albyte.html` | アルバイト勤怠打刻画面 | `src/Code.js` の Albyte 系 |
| `albyte_admin.html` | アルバイト勤怠管理者画面 | `src/Code.js` の Albyte 系 |
| `albyte_report.html` | アルバイト勤怠の月次レポート | `src/Code.js` の Albyte 系 |
| `intake.html` | 初回受付登録、電話/訪問モード、候補調整起点 | `schedule.js` と intake 系関数 |
| `intake_list.html` | 初回受付一覧、候補表示、ステータス更新 | `intakeList`, `generateTrialOpportunities`, `intakeUpdateStatus` など |
| `admin.html` | 管理者向け同意期限系ダッシュボード | `getAdminDashboard`, `runBulkActions` |
| `report.html` | AI 報告書プレビュー/生成 | 報告書生成系・保存済み報告書取得系 |
| `vacancy.html` | 空き状況表示 | `vacancyList` |

### テンプレート用途

| HTML | 用途 |
|---|---|
| `invoice_template.html` | 請求書 PDF / HTML テンプレート |
| `payroll_pdf_family.html` | 給与明細 PDF テンプレート |
| `main.js.html` | `billing.html` に埋め込まれるクライアント JS |

## スプレッドシート依存箇所

## 依存の層

### 1. メイン業務スプレッドシート

主な参照元:

- `src/Code.js`
- `src/main.gs`
- `src/get/billingGet.js`
- `src/dashboard/config.gs`
- `src/dashboard/utils/sheetUtils.js`
- `src/utility.js`

固定 ID / 既定 ID:

- `APP.SSID` in `src/Code.js`
- `FIXED_SSID` in `src/utility.js`
- `DASHBOARD_SPREADSHEET_ID` in `src/dashboard/config.gs`

これらは実質的に同じ業務台帳を指している構成です。

### 2. スケジュール専用スプレッドシート

主な参照元:

- `src/schedule.js`

固定 ID:

- `SCHEDULE_SS_ID`

## 主なシート依存

### `src/Code.js`

シート定数または初期化対象として確認できるもの:

- `施術録`
- `患者情報`
- `News`
- `フラグ`
- `予定`
- `操作ログ`
- `定型文`
- `添付索引`
- `年次確認`
- `ダッシュボード`
- `AI報告書`
- `VisitAttendance`
- `VisitAttendanceRequests`
- `VisitAttendanceStaff`
- `AlbyteStaff`
- `AlbyteAttendance`
- `AlbyteShifts`
- `PayrollEmployees`
- `PayrollGrades`
- `PayrollInsuranceStandards`
- `PayrollInsuranceOverrides`
- `PayrollRoles`
- `PayrollPayoutEvents`
- `PayrollAnnualSummaries`
- `所得税税額表`
- `申し送り`
- `Intake_Staging`

### `src/main.gs`

- `PreparedBillingMeta`
- `PreparedBillingMetaJson`
- `PreparedBillingJson`
- `請求履歴`
- `銀行情報`
- `未回収履歴`

### `src/get/billingGet.js`

- `施術録`
- `患者情報`
- `銀行情報`
- `BillingOverrides`
- `スタッフ一覧`
- `CarryOverLedger`
- `請求履歴`
- `未回収履歴`

### `src/dashboard/config.gs`

- `患者情報`
- `施術録`
- `申し送り`
- `AI報告書`
- `未回収履歴`

### `src/schedule.js`

- `Waitlist`
- `SlotMaster`
- `Assignment`
- `TrialOpportunities`
- `Config`
- `GeoCache`

## スプレッドシート依存の特徴

- 画面単位ではなく、機能単位でシート依存が混在している
- `src/Code.js` が業務シートの初期化、読み書き、UI 用整形を同時に担っている
- 請求処理は分離されつつあるが、依然として既存の `ss()` / `billingSs()` / `APP.SSID` / Script Properties に跨っている
- スケジュール系だけは別スプレッドシートに分離されている

## 外部 API / Drive / PDF / その他副作用処理

## 外部 API

### OpenAI

`src/Code.js`

- `APP.OPENAI_ENDPOINT`
- `OPENAI_API_KEY` を Script Properties から取得
- `UrlFetchApp.fetch(...)` で OpenAI Chat Completions を呼び出す
- 用途:
  - 医師向け報告書生成
  - 各種 AI レポート整形

### Chat Webhook

`src/Code.js`

- `CHAT_WEBHOOK_URL`
- `CHAT_WEBHOOK_URL_ADMIN`
- `CHAT_WEBHOOK_URL_DEFAULT`
- `UrlFetchApp.fetch(...)`
- 用途:
  - 通知送信
  - 管理者向け通知

### Geocoding

`src/schedule.js`

- `Maps.newGeocoder().geocode(addr)`
- 用途:
  - 初回受付の住所から位置情報を取得
  - 候補生成用

## Drive / Document / PDF

### 給与 PDF

`src/Code.js`

- `Utilities.newBlob(...).getAs(MimeType.PDF)`
- `folder.createFile(blob)`

### 医師向け報告書 PDF

`src/Code.js`

- `DriveApp.getFolderById(...)`
- `DriveApp.getFileById(...)`
- `DocumentApp.openById(...)`
- `copy.getAs(MimeType.PDF)`
- `doctorFolder.createFile(pdfBlob)`
- `copy.setTrashed(true)`

### HTML -> PDF 変換

`src/Code.js`

- `DocumentApp.create(...)`
- `UrlFetchApp.fetch(url).getBlob()`
- `folder.createFile(pdfBlob)`
- `DriveApp.getFileById(docId).setTrashed(true)`

### 請求関連 PDF / Drive

`src/output/billingOutput.js`

- 請求書 PDF 生成
- Drive フォルダ配下への保存
- 請求履歴への反映

### 入金結果 PDF 解析

`src/main.gs`

- `DriveApp.getFileById(fileId)`
- `file.getBlob()`
- 用途:
  - 入金結果 PDF を読み込み、請求履歴へ反映

### 申し送り画像アップロード

`src/Code.js`

- `Utilities.newBlob(...)`
- `patientFolder.createFile(blob)`
- `saved.setSharing(ANYONE_WITH_LINK, VIEW)`
- `DriveApp.getFileById(fid)`

## そのほかの副作用

### Trigger

`src/Code.js`

- `ScriptApp.newTrigger(...)`
- `ScriptApp.getProjectTriggers()`
- 勤怠同期や後処理系のトリガー管理

### Lock / Cache / Properties

主な利用箇所:

- `src/Code.js`
- `src/main.gs`
- `src/dashboard/utils/sheetUtils.js`
- `src/utility.js`

用途:

- `LockService`
  - 二重実行防止
- `CacheService`
  - ダッシュボード / 請求のキャッシュ
- `PropertiesService`
  - API key, webhook URL, spreadsheet ID, folder ID, 設定値

## 重複責務・密結合箇所

## 1. エントリーポイント重複

重複定義:

- `doGet` in `src/Code.js`
- `doGet` in `src/entrypoint.gs`
- `doGet` in `src/main.gs`

問題:

- どのルータが本番デプロイの正なのかが不明瞭
- 再構築時に「入口を変えたつもりが別の `doGet` が有効」という事故が起きやすい

## 2. スプレッドシート解決責務の重複

重複定義:

- `ss()` in `src/utility.js`
- `ss()` in `src/main.gs`
- `billingSs()` in `src/main.gs`
- `dashboardGetSpreadsheet_()` in `src/dashboard/utils/sheetUtils.js`

問題:

- 同一業務台帳へのアクセス経路が複数ある
- 固定 ID / Script Properties / Active Spreadsheet が混在している

## 3. 設定取得責務の重複

重複定義:

- `getConfig(key)` in `src/utility.js`
- `getConfig()` in `src/schedule.js`

問題:

- 同名で戻り値契約が異なる
- GAS のグローバルスコープで衝突余地がある

## 4. ダッシュボード責務の二重化

関係箇所:

- `src/Code.js`
- `src/dashboard/main.gs`
- `src/dashboard/api/*`

問題:

- 旧来ルータと新ダッシュボード分離系が共存
- `shouldHandleDashboardApi_` も複数系統に存在する

## 5. 請求責務の二重化

関係箇所:

- `src/Code.js`
- `src/main.gs`
- `src/get/billingGet.js`
- `src/logic/billingLogic.js`
- `src/output/billingOutput.js`
- `src/main.js.html`

問題:

- 分離は進んでいるが、請求 UI / 請求 API / 請求保存 / 請求 PDF / 請求履歴がまだ広く跨っている
- 互換レイヤや fallback 実装が多く、依存境界がまだ固まっていない

## 6. `src/Code.js` への機能集中

抱えている責務:

- 画面ルーティング
- 患者ヘッダ取得
- 施術録保存/更新/削除
- News
- 申し送り
- AI 報告書
- 勤怠
- 有給
- Albyte
- Payroll
- Intake 補助
- 外部通知
- Drive/PDF

問題:

- 「再構築対象を限定する」には、まずこのファイルをそのまま全置換しない前提を置く必要がある

## 7. スケジュール系の孤立

関係箇所:

- `src/schedule.js`
- `src/intake.html`
- `src/intake_list.html`

問題:

- 別スプレッドシートに依存
- 業務本体とデータ境界が異なる
- `validateSlotMaster()` が同一ファイル内で二重定義されている

## 現状の再構築ユニット候補

現状構造から見て、物理的に分けやすい単位は以下です。

1. ダッシュボード系
   - `src/dashboard/*`
   - `src/dashboard.html`
2. 請求系
   - `src/main.gs`
   - `src/get/*`
   - `src/logic/*`
   - `src/output/*`
   - `src/billing.html`
   - `src/main.js.html`
3. スケジュール候補生成系
   - `src/schedule.js`
   - `src/intake.html`
   - `src/intake_list.html`

逆に、現時点で分離が難しい塊は以下です。

1. `src/Code.js` 中の施術録 + News + 申し送り + AI報告書
2. `src/Code.js` 中の勤怠 + 有給 + Albyte + Payroll の横断部分
3. `doGet` と設定解決の基盤部分

## まとめ

再構築前提の認識として重要なのは次の 4 点です。

1. 現状は「単一アプリ」ではなく、「複数の分離途中アプリ」が同居している
2. 技術的なボトルネックは `src/Code.js` の機能集中と GAS グローバルスコープ競合
3. 再構築対象として最も限定しやすいのは「ダッシュボード系」「請求系」「スケジュール候補生成系」
4. 施術録本体とその周辺機能は密結合が強く、別扱いにしないと分析が破綻しやすい

