# treatment-log-app

社内向けの施術記録・勤怠管理アプリケーションです。Apps Script で構築しており、`doGet` から複数の HTML ビューを返しています。

## VisitAttendance（訪問スタッフ勤怠）

### Web アプリからのアクセス
- `welcome.html` のメニュー、または `?view=attendance` パラメータで勤怠ビューを開きます。
- 対象月の選択は `localStorage` に保存されるため、再訪問時は直前に見ていた月が自動で選択されます。

### 自動集計の流れ
- `syncVisitAttendance()` が施術録シートを読み込み、`VisitAttendance` シートへ自動集計を反映します。
- 承認済みの申請は `updateVisitAttendanceRequestStatus()` を通じて `VisitAttendance` に上書きされ、スタッフビューへ即時に反映されます。

### 手動同期とトリガー設定
1. スプレッドシートを開き、メニューの **勤怠管理 → 勤怠データを今すぐ同期** を実行すると、その場で `syncVisitAttendance()` が走り結果がダイアログに表示されます。
2. 同じメニューの **勤怠管理 → 日次同期トリガーを確認** で `ensureVisitAttendanceSyncTrigger()` が呼び出され、日次 0:00(JST) のトリガーが存在しない場合は自動作成されます。

### スタッフ側の修正申請
- 勤怠一覧から対象日を選び、退勤時刻・休憩時間・理由を入力して申請を送信します。
- 同一日の未処理申請が存在する場合は重複登録されないようにバリデーションしています。

### 管理者側の承認フロー
- `pending` 申請は管理者ビューに表示され、承認または差し戻し操作を行えます。
- 承認時には対象日の勤怠データが 15 分単位・18:00 までなどの制約を満たすか検証した上で `VisitAttendance` シートへ反映されます。

## デプロイ時の注意
- `Code.js` の `onOpen()` で作成されるメニューから手動同期やトリガー確認を実行してください。
- Web アプリの URL（`ScriptApp.getService().getUrl()`）を基に、`welcome.html` の各リンクが遷移します。

## 社会保険料設定
- `payroll.html` から標準報酬等級マスタや料率を管理し、従業員ごとの社会保険料プレビューを確認できます。
- `PayrollInsuranceStandards` と `PayrollInsuranceOverrides` シートにデータが保存され、標準報酬月額の自動計算・月次上書きに利用されます。
