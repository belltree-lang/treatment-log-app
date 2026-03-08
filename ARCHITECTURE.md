# Architecture

## Goal
GASベースの施術管理アプリを、モジュール分離された構造に再設計する。

目的
- UI / Application / Data を分離
- Spreadsheet依存を隔離
- 将来拡張しやすい構造にする

---

# Layers

UI Layer
HTML + Client JS

Application Layer
GAS server functions

Data Layer
Spreadsheet access

---

# Folder Structure

src/

ui/
  dashboard/
  billing/
  intake/

app/
  dashboardService.js
  billingService.js
  scheduleService.js

data/
  sheetRepository.js
  patientRepository.js
  billingRepository.js

infra/
  driveService.js
  pdfService.js

---

# Rules

UIはSpreadsheetに直接アクセスしない  
ApplicationはHtmlServiceを呼ばない  
DataはSpreadsheet操作のみ  

---

# Spreadsheet Access

Spreadsheetは必ずRepository経由

NG
SpreadsheetApp directly

OK
patientRepository.getPatient()