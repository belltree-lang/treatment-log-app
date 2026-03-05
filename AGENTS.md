# AGENTS.md

## Purpose

This repository is a Google Apps Script application for internal treatment logging, attendance management, billing, dashboard views, and related administrative workflows.

The codebase is partially modularized, but still relies heavily on Apps Script global scope. Changes should preserve behavior first and avoid introducing more cross-file ambiguity.

## Architecture

### Runtime model

- The app is deployed to Google Apps Script via `clasp`.
- Server-side code runs in Apps Script V8 and reads/writes Google Sheets, Drive, CacheService, LockService, Script Properties, and time-based triggers.
- Client-side HTML views call server functions through `google.script.run`.

### Main architectural slices

- `src/Code.js`
  - Legacy monolith.
  - Handles the treatment app router and a large amount of business logic: treatment logs, handovers, attendance, paid leave, Albyte attendance, payroll, intake, and miscellaneous utilities.
- `src/main.gs`
  - Billing-oriented entrypoint and compatibility layer.
  - Contains billing menu setup, billing cache/snapshot helpers, and billing web app behavior.
- `src/get/billingGet.js`
  - Billing data retrieval and normalization from sheets.
- `src/logic/billingLogic.js`
  - Billing calculation logic. Prefer keeping new logic here if it can be pure.
- `src/output/billingOutput.js`
  - Billing output concerns such as invoice PDF generation, billing history writes, and bank/export compatibility behavior.
- `src/dashboard/*`
  - Newer modular dashboard backend.
  - `config.gs`, `main.gs`, `api/*`, `data/*`, `auth/*`, `utils/*` separate routing, orchestration, loading, roles, and shared helpers.
- `src/*.html`
  - HTML views with inline client-side JavaScript for each feature.
- `tests/*`
  - Node-based tests that load GAS source files with `vm` and verify server-side behavior.

### Important constraints

- Apps Script uses shared global scope across server files.
- Duplicate top-level names can silently override each other.
- There are multiple historical entrypoint patterns in this repo. Do not add new top-level routing functions casually.
- Spreadsheet and Drive IDs are currently hardcoded in several places. Avoid spreading that pattern further.

## Coding Rules

### General

- Preserve existing production behavior unless the task explicitly requires behavioral change.
- Make the smallest coherent change that fixes the problem.
- Prefer extraction and reuse over copying logic into another file.
- Keep all edits ASCII unless the file already uses Japanese labels or non-ASCII content that must be preserved.

### Apps Script server-side rules

- Avoid introducing new top-level function names that are vague or likely to collide.
- Before adding a shared helper, check whether an equivalent helper already exists in:
  - `src/Code.js`
  - `src/main.gs`
  - `src/get/billingGet.js`
  - `src/dashboard/utils/sheetUtils.js`
- Prefer pure functions for:
  - billing calculations
  - normalization
  - formatting
  - row-to-object mapping
- Keep sheet I/O at the edges. Separate:
  - sheet reads/writes
  - business rules
  - response shaping for the UI
- If a function depends on sheet headers, make header handling explicit and defensive.
- When touching caching code, preserve cache key isolation by month/user/role where relevant.
- When touching triggers or batch jobs, preserve locking and idempotency behavior.

### HTML / client-side rules

- Do not move large new business rules into HTML inline scripts unless there is no better existing location.
- Prefer server-side validation even if the UI also validates.
- Keep DOM IDs and `google.script.run` contracts stable unless all call sites are updated together.
- Avoid mixing unrelated feature changes into the same HTML file edit.

### Architecture rules for new code

- Billing changes should follow the current layering when possible:
  - data loading in `src/get/billingGet.js`
  - calculation logic in `src/logic/billingLogic.js`
  - output/write concerns in `src/output/billingOutput.js`
  - UI glue in `src/main.gs` or `src/main.js.html`
- Dashboard changes should stay within `src/dashboard/*` unless they must integrate with legacy treatment code.
- Do not add more unrelated logic into `src/Code.js` if an existing modular location already fits.
- If a change must touch `src/Code.js`, isolate it to the smallest possible region and avoid broad refactors unless requested.

### Configuration rules

- Prefer centralizing new configuration behind Script Properties or existing config helpers.
- Do not hardcode new spreadsheet IDs, folder IDs, or admin emails without a strong reason.
- If you must add a constant, place it near the relevant feature’s existing constants.

## Commit Style

- Use small, reviewable commits.
- Commit messages should be imperative and scoped.
- Preferred format:
  - `billing: fix prepared month validation`
  - `dashboard: isolate staff scope cache key`
  - `attendance: guard duplicate paid leave requests`
  - `docs: add repo agent guidance`
- Avoid vague messages such as:
  - `fix stuff`
  - `update code`
  - `changes`
- Do not mix unrelated domains in one commit unless the change is genuinely cross-cutting.

## Testing Rules

### Before changing code

- Read the relevant tests first when modifying:
  - billing behavior
  - dashboard behavior
  - treatment record save/delete behavior
  - attendance or paid leave behavior

### After changing code

- Run the smallest relevant test set first.
- If the touched area already has focused tests, run those tests before considering broader coverage.
- If you change shared billing logic, also run neighboring billing regression tests.
- If you change dashboard orchestration or routing, run dashboard routing and data tests.

### Test expectations

- Add or update tests for behavior changes, bug fixes, and regressions when practical.
- Prefer focused regression tests over broad rewrites of existing test fixtures.
- Keep tests deterministic. Avoid real network calls or live Apps Script dependencies.
- Follow the existing pattern: tests load GAS files with Node `vm` and stub the Apps Script runtime.

### Current repo reality

- The repository contains tests under `tests/`, but no checked-in `package.json` or runner config was found at the root.
- If you add or update tests, document or preserve the expected execution method used by the repo’s maintainers.
- If you cannot run tests in the current environment, say so explicitly in your summary.

## Change Review Checklist

- Does this introduce or rename a top-level GAS function?
- Could this collide with another helper in global scope?
- Is spreadsheet access isolated from calculation logic where possible?
- Are UI and server contracts still aligned?
- Are cache keys, month keys, and patient ID normalization still consistent?
- Are tests updated for the changed behavior?

