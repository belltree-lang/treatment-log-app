# Implementation Rules

このドキュメントは REBUILD_SPEC.md を実装する際の制約条件を定義する。

## 1. 既存コード

既存コードは legacy として扱う。

変更禁止。

src/legacy/*
には一切変更を加えない。

## 2. 新規コード

新規コードは以下に作成する。

src/rebuild/

## 3. Spreadsheet

既存 Spreadsheet 構造は変更禁止。

新しいシート追加禁止。

列追加禁止。

既存列のみ使用。

## 4. API

既存 GAS 関数は変更禁止。

必要な場合は wrapper を作成する。

## 5. UI

UI は rebuild ディレクトリ内に新規作成する。

既存 HTML は変更禁止。

## 6. 副作用

Drive
PDF
External API

これらは既存関数を呼び出すのみ。

新規実装は禁止。

## 7. スコープ

実装対象

Dashboard
Billing
ScheduleCandidateGenerator

非対象

TreatmentLog
Attendance
Payroll
Albyte