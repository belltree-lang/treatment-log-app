/**
 * ダッシュボード共通の設定値を集約する。
 */
const DASHBOARD_PREFIX = 'DASHBOARD_';
const DASHBOARD_CACHE_TTL_SECONDS = 60;
const DATE_FORMAT = 'yyyy/MM/dd';
const DEBUG_MODE = false;
const DEFAULT_TZ = 'Asia/Tokyo';
// ダッシュボードが参照するスプレッドシートと請求書フォルダ
const DASHBOARD_SPREADSHEET_ID = '1ajnW9Fuvu0YzUUkfTmw0CrbhrM3lM5tt5OA1dK2_CoQ';
const DASHBOARD_INVOICE_FOLDER_ID = '1EG-GB3PbaUr9C1LJWlaf_idqoYF-19Ux';
// ダッシュボードが参照する各シート名
const DASHBOARD_SHEET_PATIENTS = '患者情報';
const DASHBOARD_SHEET_TREATMENTS = '施術録';
const DASHBOARD_SHEET_NOTES = '申し送り';
const DASHBOARD_SHEET_AI_REPORTS = 'AI報告書';
const DASHBOARD_SHEET_UNPAID_HISTORY = '未回収履歴';

// 未回収アラート判定に利用する連続月数のデフォルト値
const DASHBOARD_UNPAID_ALERT_MONTHS = 3;
