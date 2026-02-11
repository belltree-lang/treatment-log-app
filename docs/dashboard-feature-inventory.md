# ダッシュボード画面 機能棚卸し

## ■ 画面構成一覧（上から順に）
1. ヘッダー
   - タイトル（施術者ダッシュボード）
   - ログインユーザー表示
   - データ生成時刻表示
   - 「最新の情報を取得」ボタン
2. エラーボックス
3. 概要カード（3列）
   - ①請求
   - ②同意
   - ③施術実績
4. 未回収アラート
5. 今日・昨日の訪問（タイムライン）
6. 担当患者一覧（アコーディオン）
7. ローディングオーバーレイ

## ■ 各セクションの目的
- ヘッダー: 利用者コンテキストと手動リフレッシュ操作を提供。
- エラーボックス: API失敗・メタエラー・遷移URL未設定などを通知。
- 概要カード:
  - ①請求: 前月請求未確認患者の一覧を即時把握。
  - ②同意: 同意期限関連・同意取得/通院日関連の不備把握。
  - ③施術実績: 今日件数と直近1日件数の2指標表示。
- 未回収アラート: 連続未回収（月次）の患者を金額付きで警告。
- 今日・昨日の訪問: 直近訪問履歴と報告書作成ヒント有無の確認。
- 担当患者一覧: 個別患者の状況（同意・AI報告・請求書・申し送り）確認と既読化。
- ローディング: 取得中の操作抑止と待機表示。

## ■ 表示データの取得元
- 画面データ本体は `getDashboardData` の単一レスポンス。
- `getDashboardData` 内部で以下を集約:
  - 患者情報: `loadPatientInfo`（患者情報シート）
  - 申し送り: `loadNotes`（申し送りシート + ScriptProperties 既読）
  - AI報告: `loadAIReports`（AI報告書シート）
  - 請求書: `loadInvoices`（請求書フォルダ + 患者マスタ紐付け）
  - 施術ログ: `loadTreatmentLogs`（施術録シート）
  - 担当者: `assignResponsibleStaff`（施術ログから前月末最終施術者）
  - 未回収: `loadUnpaidAlerts`（未回収履歴シート）
  - タスク: `getTasks`（同意期限/申し送り遅延/AI報告遅延/請求未確認）
  - 訪問: `getTodayVisits`（今日・昨日の施術抽出）
- 設定上の主データソース:
  - Spreadsheet: `DASHBOARD_SPREADSHEET_ID`
  - Invoice folder: `DASHBOARD_INVOICE_FOLDER_ID`
  - シート名: 患者情報 / 施術録 / 申し送り / AI報告書 / 未回収履歴

## ■ クリック可能要素一覧
- 「最新の情報を取得」ボタン（再取得）
- 概要カード ①請求/②同意 の患者行（overview-row）
- 未回収アラート内の患者名リンク
- 今日・昨日の訪問の患者名リンク
- 担当患者一覧の summary 行（患者IDがあり遷移URL有効時）
- 担当患者一覧の請求書PDFリンク
- 担当患者アコーディオン開閉（details toggle）

## ■ 各ボタンの遷移先
- 最新の情報を取得: 画面内再描画（遷移なし）
- 概要カード患者行: `.../exec?view=record&id=<patientId>` を新規タブ
- 未回収アラート患者名: 同上（新規タブ）
- 訪問タイムライン患者名: 同上（新規タブ）
- 担当患者 summary: 同上（新規タブ）
- 請求書PDFを開く: `invoiceUrl`（Drive PDF想定）を新規タブ

## ■ API一覧
### HTTP (doGet)
- `GET ?action=getDashboardData` または `/getDashboardData`
  - 応答: ダッシュボードJSON（tasks, todayVisits, patients, unpaidAlerts, warnings, overview, meta）
- `GET ?view=dashboard` または `/dashboard`
  - 応答: ダッシュボードHTML

### google.script.run（HTML→GAS）
- `getDashboardData(args)`
- `markAsRead({ patientId })`

### 関連内部API（集約用）
- `getTasks(options)`
- `getTodayVisits(options)`
- `loadUnpaidAlerts(options)`
- `loadPatientInfo/loadNotes/loadAIReports/loadInvoices/loadTreatmentLogs`

## ■ 未使用UI要素
- CSS定義のみでDOM生成されない要素
  - `.warning-list`
  - `.warning-item`
- 関数定義のみで現在の描画経路で未使用
  - `formatTaskLabel`
  - `formatDateLabel`
- 取得はされるが画面表示されないデータ
  - `tasks`（ログ用途・overview材料としては利用）
  - `warnings`（収集されるがUI表示なし）
- 一時抑止中のUI
  - `DASHBOARD_SUPPRESS_HANDOVER_REMINDER_UI = true` により「報告書作成ヒント未入力」表示を抑止

## ■ 技術的負債候補
1. 単一HTML内script肥大化
   - 描画・状態管理・遷移制御・API呼び出しが1ファイル集中で保守コスト高。
2. dead code / 死蔵スタイル
   - 未使用関数・未使用CSSが残存し意図不明瞭。
3. suppressフラグ常時ON運用
   - handover関連表示をコード上保持しつつ常時無効化（恒久化の危険）。
4. エラーハンドリングの表示粒度
   - warningsは収集されるがユーザー非表示で運用気付きにくい。
5. getDashboardData責務過大
   - データ取得/フィルタ/集約/overview生成まで単一関数に集中。
6. ルーティング方式の二重化
   - `google.script.run` と `fetch(?action=...)` の二経路維持で挙動差分リスク。
7. URL依存の防御がUI側中心
   - 遷移URL未設定時の抑止はあるが運用設定不備を事前検知しづらい。
