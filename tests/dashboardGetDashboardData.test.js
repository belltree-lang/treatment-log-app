const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const sheetUtilsCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'utils', 'sheetUtils.js'),
  'utf8'
);
const configCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'config.gs'),
  'utf8'
);
const dashboardCode = fs.readFileSync(
  path.join(__dirname, '..', 'src', 'dashboard', 'api', 'getDashboardData.js'),
  'utf8'
);

function createContext(overrides = {}) {
  const context = {
    console,
    JSON,
    Date,
    Set,
    Utilities: {
      formatDate: (date, _tz, _fmt) => date.toISOString()
    },
    Session: {
      getScriptTimeZone: () => 'Asia/Tokyo',
      getActiveUser: () => ({ getEmail: () => overrides.activeEmail || 'session@example.com' })
    },
    calculateConsentExpiry_: value => {
      const base = new Date(value);
      if (Number.isNaN(base.getTime())) return '';
      const monthsToAdd = base.getDate() <= 15 ? 5 : 6;
      const target = new Date(base);
      target.setMonth(target.getMonth() + monthsToAdd + 1, 0);
      return target.toISOString().slice(0, 10);
    }
  };
  Object.assign(context, overrides);
  vm.createContext(context);
  vm.runInContext(configCode, context);
  vm.runInContext(sheetUtilsCode, context);
  vm.runInContext(dashboardCode, context);
  return context;
}

function testAggregatesDashboardData() {
  const patientInfo = {
    patients: { '001': { name: '山田太郎', raw: { '同意年月日': '2024-07-16' } } },
    warnings: ['p1']
  };
  const notes = {
    notes: { '001': { preview: '最新メモ', when: '2025-02-01', unread: true, authorEmail: 'note@example.com', lastReadAt: '2025-01-01T00:00:00Z', row: 5 } },
    warnings: ['n1']
  };
  const aiReports = { reports: { '001': '2025-01-15 10:00' }, warnings: ['a1'] };
  const invoices = { invoices: { '001': 'https://example.com/invoice.pdf' }, warnings: ['i1'] };
  const treatmentLogs = { logs: [{ patientId: '001', timestamp: new Date('2025-02-01T09:00:00Z'), dateKey: '2025-02-01', createdByEmail: 'user@example.com', staffKeys: { email: 'user@example.com', name: '', staffId: '' } }], warnings: ['t1'] };
  const responsible = { responsible: { '001': 'staff@example.com' }, warnings: ['r1'] };
  const tasksResult = { tasks: [{ type: 'consentWarning', patientId: '001' }], warnings: ['task'] };
  const visitsResult = { visits: [{ patientId: '001', dateKey: '2025-02-01', time: '10:00' }], warnings: ['visit'] };
  const unpaidAlerts = { alerts: [{ patientId: '001', patientName: '山田太郎', consecutiveMonths: 3, totalAmount: 15000, months: [], followUp: { phone: false, visit: false } }], warnings: ['u1'] };

  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T12:00:00Z'),
    patientInfo,
    notes,
    aiReports,
    invoices,
    treatmentLogs,
    responsible,
    unpaidAlerts,
    tasksResult,
    visitsResult
  });

  assert.strictEqual(result.meta.user, 'user@example.com');
  assert.ok(result.meta.generatedAt, 'generatedAt should be present');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.tasks)), []);
  assert.deepStrictEqual(result.todayVisits, visitsResult.visits);
  assert.strictEqual(result.patients.length, 1);
  const normalizedPatient = JSON.parse(JSON.stringify(result.patients[0]));
  assert.deepStrictEqual(normalizedPatient, {
    patientId: '001',
    name: '山田太郎',
    consentExpiry: '2025-01-31T00:00:00.000Z',
    responsible: 'staff@example.com',
    invoiceUrl: 'https://example.com/invoice.pdf',
    aiReportAt: '2025-01-15 10:00',
    note: {
      patientId: '001',
      preview: '最新メモ',
      when: '2025-02-01',
      unread: true,
      lastReadAt: '2025-01-01T00:00:00Z',
      authorEmail: 'note@example.com',
      row: 5
    },
    statusTags: [
      { type: 'consent', label: '⚠ 期限超過', priority: 'high' },
      { type: 'report', label: '作成済' }
    ]
  });
  assert.strictEqual(result.unpaidAlerts.length, 1, '未回収アラートが伝搬する');
  assert.strictEqual(result.unpaidAlerts[0].patientId, '001');
  const warnings = JSON.parse(JSON.stringify(result.warnings)).sort();
  assert.deepStrictEqual(warnings, ['a1', 'i1', 'n1', 'p1', 'r1', 't1', 'u1', 'visit'].sort());
}



function testPatientStatusTagsGeneration() {
  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '期限内未取得', raw: { '同意年月日': '2024-08-16' } },
        '002': { name: '期限超過未取得', raw: { '同意年月日': '2024-07-16' } },
        '003': { name: '報告書作成済', raw: { '同意年月日': '2024-08-16' } },
        '004': { name: '同意取得確認済', raw: { '同意年月日': '2024-08-16', '同意書取得確認': '済' } },
        '005': { name: '同意日更新後', raw: { '同意年月日': '2024-09-16' } },
        '006': { name: '期限迫る未取得', raw: { '同意年月日': '2024-08-16' } }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: {
      reports: {
        '003': '2025-01-30'
      },
      warnings: []
    },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  const patientsById = {};
  result.patients.forEach(entry => {
    patientsById[entry.patientId] = JSON.parse(JSON.stringify(entry.statusTags));
  });

  assert.deepStrictEqual(patientsById['001'], [
    { type: 'consent', label: '📄 要対応', priority: 'low' },
    { type: 'report', label: '未作成' }
  ], '1. 期限内未取得は要対応＋未作成を表示する');

  assert.deepStrictEqual(patientsById['002'], [
    { type: 'consent', label: '⚠ 期限超過', priority: 'high' },
    { type: 'report', label: '未作成' }
  ], '2. 期限超過未取得は期限超過＋未作成を表示する');

  assert.deepStrictEqual(patientsById['003'], [
    { type: 'consent', label: '📄 要対応', priority: 'low' },
    { type: 'report', label: '作成済' }
  ], '3. 報告書作成済は要対応＋作成済を表示する');

  assert.deepStrictEqual(patientsById['004'], [], '4. 同意取得確認ありは両タグを非表示にする');

  assert.deepStrictEqual(patientsById['005'], [
    { type: 'report', label: '未作成' }
  ], '5. 同意期限30日超はconsent非表示だがreportは表示する');

  assert.deepStrictEqual(patientsById['006'], [
    { type: 'consent', label: '📄 要対応', priority: 'low' },
    { type: 'report', label: '未作成' }
  ], '6. 同意年月日から算出した期限内は要対応を表示する');
}

function testConsentOverviewMatchesPatientStatusTags() {
  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '期限内未取得', raw: { '同意年月日': '2024-08-16' } },
        '002': { name: '期限超過未取得', raw: { '同意年月日': '2024-07-16' } },
        '003': { name: '同意取得確認済', raw: { '同意年月日': '2024-08-16', '同意書取得確認': '済' } },
        '004': { name: '期限未登録', raw: {} },
        '005': { name: '期限迫る未取得', raw: { '同意年月日': '2024-08-16' } }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: {
      tasks: [
        { patientId: '003', type: 'consentExpired', detail: 'legacy' },
        { patientId: '004', type: 'consentWarning', detail: 'legacy' }
      ],
      warnings: []
    },
    visitsResult: { visits: [], warnings: [] }
  });

  const overviewItems = JSON.parse(JSON.stringify(result.overview.consentRelated.items));
  assert.deepStrictEqual(overviewItems, [
    { patientId: '002', name: '期限超過未取得', subText: '⚠ 同意期限超過（1日超過）' },
    { patientId: '001', name: '期限内未取得', subText: '同意期限（残27日）' },
    { patientId: '005', name: '期限迫る未取得', subText: '同意期限（残27日）' }
  ], '上段同意ブロックは同意年月日から算出した期限と同意書取得確認で表示する');

  const patientsById = {};
  result.patients.forEach(entry => {
    patientsById[entry.patientId] = JSON.parse(JSON.stringify(entry.statusTags));
  });

  assert.deepStrictEqual(patientsById['001'].filter(tag => tag.type === 'consent'), [{ type: 'consent', label: '📄 要対応', priority: 'low' }], 'Case1: 期限内・未取得');
  assert.deepStrictEqual(patientsById['002'].filter(tag => tag.type === 'consent'), [{ type: 'consent', label: '⚠ 期限超過', priority: 'high' }], 'Case2: 期限超過・未取得');
  assert.deepStrictEqual((patientsById['003'] || []).filter(tag => tag.type === 'consent'), [], 'Case3: 同意取得確認済');
  assert.deepStrictEqual((patientsById['004'] || []).filter(tag => tag.type === 'consent'), [], 'Case4: 期限未登録');
  assert.deepStrictEqual((patientsById['005'] || []).filter(tag => tag.type === 'consent'), [{ type: 'consent', label: '📄 要対応', priority: 'low' }], 'Case5: 期限内・未取得');
}


function testEvaluateConsentStatusBoundaries() {
  const ctx = createContext();
  const today = new Date('2025-02-01T00:00:00Z');

  const expired = JSON.parse(JSON.stringify(ctx.evaluateConsentStatus_(new Date('2025-01-31T00:00:00Z'), today)));
  assert.deepStrictEqual(expired, { type: 'expired', days: -1, priority: 'high' }, 'expired: 期限超過は high');

  const warning = JSON.parse(JSON.stringify(ctx.evaluateConsentStatus_(new Date('2025-02-15T00:00:00Z'), today)));
  assert.deepStrictEqual(warning, { type: 'warning', days: 14, priority: 'medium' }, 'warning: 14日以内は medium');

  const normal = JSON.parse(JSON.stringify(ctx.evaluateConsentStatus_(new Date('2025-02-16T00:00:00Z'), today)));
  assert.deepStrictEqual(normal, { type: 'normal', days: 15, priority: 'low' }, 'normal: 15〜30日は low');

  const hiddenOver30 = ctx.evaluateConsentStatus_(new Date('2025-03-04T00:00:00Z'), today);
  assert.strictEqual(hiddenOver30, null, '30日超は非表示');

  const hiddenNoDate = ctx.evaluateConsentStatus_(null, today);
  assert.strictEqual(hiddenNoDate, null, '同意年月日なしは非表示');
}


function testConsentOver30DaysStillShowsReportTag() {
  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '同意期限30日超', raw: { '同意年月日': '2024-09-16' } }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  const tags = JSON.parse(JSON.stringify(result.patients[0].statusTags));
  assert.deepStrictEqual(tags.filter(tag => tag.type === 'consent'), [], '30日超はconsentタグ非表示');
  assert.deepStrictEqual(tags.filter(tag => tag.type === 'report'), [{ type: 'report', label: '未作成' }], '30日超でもreportタグは表示');
}

function testEvaluateConsentStatusIsTimeIndependentForSameDate() {
  const ctx = createContext();

  const onDeadline = JSON.parse(JSON.stringify(ctx.evaluateConsentStatus_(new Date('2025-02-01T00:00:00Z'), new Date('2025-02-01T12:34:56Z'))));
  assert.deepStrictEqual(onDeadline, { type: 'warning', days: 0, priority: 'medium' }, '期限当日(diffDays=0)はwarning');

  const dayBeforeLateNight = JSON.parse(JSON.stringify(ctx.evaluateConsentStatus_(new Date('2025-02-01T00:00:00Z'), new Date('2025-01-31T23:59:00Z'))));
  assert.deepStrictEqual(dayBeforeLateNight, { type: 'warning', days: 1, priority: 'medium' }, '期限前日23:59はwarning');

  const sameDayAfterMidnight = JSON.parse(JSON.stringify(ctx.evaluateConsentStatus_(new Date('2025-02-01T00:00:00Z'), new Date('2025-02-01T00:01:00Z'))));
  assert.deepStrictEqual(sameDayAfterMidnight, { type: 'warning', days: 0, priority: 'medium' }, '期限当日00:01はwarning（時刻非依存）');
}

function testConsentAcquiredJudgmentHandlesFalseyStringsConsistently() {
  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '未取得', raw: { '同意年月日': '2024-08-16', '同意書取得確認': '未取得' } },
        '002': { name: 'FALSE文字列', raw: { '同意年月日': '2024-08-16', '同意書取得確認': 'FALSE' } },
        '003': { name: 'ゼロ文字列', raw: { '同意年月日': '2024-08-16', '同意書取得確認': '0' } },
        '004': { name: '取得済み', raw: { '同意年月日': '2024-08-16', '同意書取得確認': '済' } },
        '005': { name: 'boolean true', raw: { '同意年月日': '2024-08-16', '同意書取得確認': true } }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  const overviewPatientIds = JSON.parse(JSON.stringify(result.overview.consentRelated.items)).map(item => item.patientId);
  assert.deepStrictEqual(overviewPatientIds, ['002', '003', '001'], '未取得/FALSE/0 は上段同意に表示する');

  const patientsById = {};
  result.patients.forEach(entry => {
    patientsById[entry.patientId] = JSON.parse(JSON.stringify(entry.statusTags));
  });

  ['001', '002', '003'].forEach(patientId => {
    assert.deepStrictEqual(
      (patientsById[patientId] || []).filter(tag => tag.type === 'consent'),
      [{ type: 'consent', label: '📄 要対応', priority: 'low' }],
      `下段同意タグにも表示される: ${patientId}`
    );
  });

  ['004', '005'].forEach(patientId => {
    assert.deepStrictEqual(
      (patientsById[patientId] || []).filter(tag => tag.type === 'consent'),
      [],
      `済/true は下段同意タグを表示しない: ${patientId}`
    );
    assert.strictEqual(overviewPatientIds.includes(patientId), false, `済/true は上段同意に表示しない: ${patientId}`);
  });
}



function testParseJapaneseEraDateAndResolveConsentExpiry() {
  const ctx = createContext();

  assert.strictEqual(ctx.parseJapaneseEraDate_('令和7年6月26日'), '2025-06-26', '令和を西暦に変換する');
  assert.strictEqual(ctx.parseJapaneseEraDate_('令和元年5月1日'), '2019-05-01', '令和元年を1年として扱う');
  assert.strictEqual(ctx.parseJapaneseEraDate_('平成30年4月10日'), '2018-04-10', '平成年号を西暦に変換する');
  assert.strictEqual(ctx.parseJapaneseEraDate_('不正文字列'), null, '不正文字列は null を返す');

  const reiwaResolved = JSON.parse(JSON.stringify(ctx.resolveConsentExpiry_({
    raw: { '同意年月日': '令和7年6月26日' }
  })));
  assert.strictEqual(reiwaResolved.value, '2025-12-31T00:00:00.000Z', '令和7年6月26日は月末期限を算出できる');

  const reiwaSlashResolved = JSON.parse(JSON.stringify(ctx.resolveConsentExpiry_({
    raw: { '同意年月日': '令和7/6/26' }
  })));
  assert.strictEqual(reiwaSlashResolved.value, '2025-12-31T00:00:00.000Z', '令和7/6/26 も月末期限を算出できる');

  const heiseiResolved = JSON.parse(JSON.stringify(ctx.resolveConsentExpiry_({
    raw: { '同意年月日': '平成30年4月10日' }
  })));
  assert.strictEqual(heiseiResolved.value, '2018-09-30T00:00:00.000Z', '平成30年4月10日を変換して期限算出できる');

  const invalidResolved = JSON.parse(JSON.stringify(ctx.resolveConsentExpiry_({
    raw: { '同意年月日': '不正文字列' }
  })));
  assert.strictEqual(invalidResolved.value, null, '不正文字列は同意期限を解決しない');
}

function testConsentDateParsingFormatsAndResolverPriority() {
  const ctx = createContext();
  const now = new Date('2025-02-01T00:00:00Z');

  const resolved = ctx.resolveConsentExpiry_({
    raw: {
      '同意年月日': '2024-08-16'
    }
  });
  const normalizedResolved = JSON.parse(JSON.stringify(resolved));
  assert.strictEqual(normalizedResolved.source, "raw['同意年月日']");
  assert.strictEqual(normalizedResolved.value, '2025-02-28T00:00:00.000Z', '同意年月日から同意期限を算出する');

  const resolvedWithoutConsentDate = ctx.resolveConsentExpiry_({ raw: {} });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(resolvedWithoutConsentDate)), {
    value: null,
    source: ''
  }, '同意年月日がなければ null を返す');

  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now,
    patientInfo: {
      patients: {
        '001': { name: 'A-hyphen', raw: { '同意年月日': '2024-08-16' } },
        '002': { name: 'B-slash', raw: { '同意年月日': '2024/08/16' } },
        '003': { name: 'C-japanese', raw: { '同意年月日': '2024年8月16日' } },
        '004': { name: 'D-iso', raw: { '同意年月日': '2024-08-16T00:00:00Z' } },
        '005': { name: 'E-date', raw: { '同意年月日': new Date('2024-08-16T00:00:00Z') } },
        '006': { name: 'F-invalid', raw: { '同意年月日': 'invalid-date' } },
        '007': { name: 'G-acquired', raw: { '同意年月日': '2024-08-16', '同意書取得確認': '済' } },
        '008': { name: 'H-no-consent-date', raw: {} }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  const overviewIds = JSON.parse(JSON.stringify(result.overview.consentRelated.items)).map(item => item.patientId);
  assert.deepStrictEqual(overviewIds, ['001', '002', '003', '004', '005'], '和暦を含む解釈可能な同意年月日の患者のみ表示される');

  const patientsById = {};
  result.patients.forEach(entry => {
    patientsById[entry.patientId] = JSON.parse(JSON.stringify(entry.statusTags));
  });
  ['001', '002', '003', '004', '005'].forEach(patientId => {
    assert.deepStrictEqual((patientsById[patientId] || []).filter(tag => tag.type === 'consent'), [{ type: 'consent', label: '📄 要対応', priority: 'low' }], `同意タグが表示される: ${patientId}`);
  });
    assert.deepStrictEqual((patientsById['006'] || []).filter(tag => tag.type === 'consent'), [], '不正文字列は同意タグの表示対象外');
  assert.deepStrictEqual((patientsById['007'] || []).filter(tag => tag.type === 'consent'), [], '取得済み判定は従来通り同意タグ非表示');
  assert.deepStrictEqual((patientsById['008'] || []).filter(tag => tag.type === 'consent'), [], '同意年月日なしでは同意タグを表示しない');
}

function testConsentDateParseFailureCanBeDebugLogged() {
  const logs = [];
  const ctx = createContext({ DASHBOARD_DEBUG_CONSENT: true });
  ctx.dashboardLogContext_ = (label, details) => {
    logs.push({ label, details: String(details || '') });
  };

  ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: 'invalid-overview', raw: { '同意年月日': 'invalid-date' } }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  const parseFailureLogs = logs.filter(entry => entry.label === 'consent-date-parse-failed');
  assert.ok(parseFailureLogs.length >= 1, '同意年月日の形式不正時は parse 失敗ログを出す');
  assert.ok(parseFailureLogs.some(entry => entry.details.includes('\"raw\":\"invalid-date\"')), 'parse 失敗ログに入力値を含める');
}

function testStaffMatchingUsesEmailNameAndStaffIdWithLogs() {
  const logEntries = [];
  const ctx = createContext();
  ctx.dashboardLogContext_ = (label, details) => {
    logEntries.push({ label, details: String(details || '') });
  };

  const resultByEmail = ctx.getDashboardData({
    user: 'belltree@belltree1102.com',
    patientInfo: { patients: { '001': { name: '患者A' }, '002': { name: '患者B' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { row: 2, patientId: '001', createdByEmail: 'belltree+billing@belltree1102.com', staffName: '別名', staffId: '', staffKeys: { email: 'belltree+billing@belltree1102.com', name: '別名', staffId: '' } },
        { row: 3, patientId: '002', createdByEmail: '', staffName: '管理者ID', staffId: 'admin001', staffKeys: { email: '', name: '管理者ID', staffId: 'admin001' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(resultByEmail.meta.error, undefined);

  const resultByStaffId = ctx.getDashboardData({
    user: 'admin001',
    patientInfo: { patients: { '001': { name: '患者A' }, '002': { name: '患者B' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { row: 2, patientId: '001', createdByEmail: 'other@example.com', staffName: '別名', staffId: '', staffKeys: { email: 'other@example.com', name: '別名', staffId: '' } },
        { row: 3, patientId: '002', createdByEmail: '', staffName: '管理者ID', staffId: 'admin001', staffKeys: { email: '', name: '管理者ID', staffId: 'admin001' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(resultByStaffId.meta.error, undefined);

  const strategyLog = logEntries.find(entry => entry.label === 'getDashboardData:staffMatchStrategy');
  const matchedCountLog = logEntries.find(entry => entry.label === 'getDashboardData:staffMatchedLogs');
  const matchedIdsLog = logEntries.find(entry => entry.label === 'getDashboardData:matchedPatientIds');
  assert.ok(strategyLog, 'staffMatchStrategy ログが出力される');
  assert.ok(matchedCountLog, 'staffMatchedLogs ログが出力される');
  assert.ok(matchedIdsLog, 'matchedPatientIds ログが出力される');
  assert.ok(Number(matchedCountLog.details) >= 1, '少なくとも1件の一致ログがある');
}



function testStaffConsentScopeMetricsAreLogged() {
  const logEntries = [];
  const ctx = createContext();
  ctx.dashboardLogContext_ = (label, details) => {
    logEntries.push({ label, details: String(details || '') });
  };

  const result = ctx.getDashboardData({
    user: 'staff@example.com',
    now: new Date('2025-02-20T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '患者A', raw: { '同意年月日': '2024-09-16' } },
        '002': { name: '患者B', raw: { '同意年月日': '2024-09-16' } },
        '003': { name: '患者C', raw: { '同意年月日': '2024-09-16', '同意書取得確認': '済' } },
        '004': { name: '患者D', raw: {} }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { patientId: '001', timestamp: new Date('2025-02-10T09:00:00Z'), staffKeys: { email: 'staff@example.com', name: '', staffId: '' } },
        { patientId: '002', timestamp: new Date('2024-12-01T09:00:00Z'), staffKeys: { email: 'staff@example.com', name: '', staffId: '' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(result.meta.error, undefined);

  const scopeMetricsLog = logEntries.find(entry => entry.label === 'getDashboardData:consentScopeMetrics');
  assert.ok(scopeMetricsLog, 'consentScopeMetrics ログが出力される');
  const metrics = JSON.parse(scopeMetricsLog.details);
  assert.strictEqual(metrics.totalPatients, 4);
  assert.strictEqual(metrics.consentEligiblePatients, 2);
  assert.strictEqual(metrics.visiblePatientIdsSize, 1);
  assert.strictEqual(metrics.consentEligibleButOutOfScope, 1);

  const missingByRecentLog = logEntries.find(entry => entry.label === 'getDashboardData:consentMissingByRecentLog');
  assert.ok(missingByRecentLog, 'consentMissingByRecentLog ログが出力される');
  assert.ok(missingByRecentLog.details.indexOf('=1') >= 0, '直近50日ログなし件数がログに含まれる');

  assert.ok(logEntries.some(entry => entry.label === 'getDashboardData:visibleScopeRoutes'), 'visibleScopeRoutes ログが出力される');
  assert.ok(logEntries.some(entry => entry.label === 'buildDashboardPatients_:scope'), 'buildDashboardPatients_:scope ログが出力される');
  assert.ok(logEntries.some(entry => entry.label === 'buildOverviewFromConsent_:scope'), 'buildOverviewFromConsent_:scope ログが出力される');
}

function testStaffConsentEligibilityEvaluatesOnlyVisibleOrMatchedPatients() {
  const logEntries = [];
  const ctx = createContext();
  ctx.dashboardLogContext_ = (label, details) => {
    logEntries.push({ label, details: String(details || '') });
  };

  const result = ctx.getDashboardData({
    user: 'staff@example.com',
    now: new Date('2025-02-20T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '患者A', raw: { '同意年月日': '2024-09-16' } },
        '002': { name: '患者B', raw: { '同意年月日': '2024-09-16' } },
        '003': { name: '患者C', raw: { '同意年月日': '2024-09-16' } }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { patientId: '001', timestamp: new Date('2025-02-10T09:00:00Z'), staffKeys: { email: 'staff@example.com', name: '', staffId: '' } },
        { patientId: '002', timestamp: new Date('2024-12-01T09:00:00Z'), staffKeys: { email: 'staff@example.com', name: '', staffId: '' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(result.meta.error, undefined);

  const eligibilityLogs = logEntries
    .filter(entry => entry.label === 'getDashboardData:consentEligibilityPatient')
    .map(entry => JSON.parse(entry.details).pid)
    .sort();
  assert.deepStrictEqual(eligibilityLogs, ['001'], '可視範囲内の患者のみ同意評価ログに含まれる');

  const consentDebugLogs = logEntries
    .filter(entry => entry.label === 'consent-eligible-debug')
    .map(entry => JSON.parse(entry.details).pid)
    .sort();
  assert.deepStrictEqual(consentDebugLogs, ['001'], '可視範囲外の患者は一致していても同意デバッグログを出力しない');
}

function testVisitSummaryWhenTodayIsZeroUsesLatestPastDayCount() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        if (fmt === 'yyyy-MM-dd') return new Date(date).toISOString().slice(0, 10);
        return new Date(date).toISOString();
      }
    }
  });

  const result = ctx.buildOverviewFromTreatmentProgress_(
    [
      { patientId: '001', dateKey: '2025-01-31', time: '09:00' }
    ],
    new Date('2025-02-01T00:00:00Z'),
    'Asia/Tokyo'
  );

  assert.deepStrictEqual(JSON.parse(JSON.stringify(result)), {
    todayCount: 0,
    latestDayCount: 1
  });
}

function testVisitSummaryWhenTodayHasTwoUsesTodayCountForBoth() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        if (fmt === 'yyyy-MM-dd') return new Date(date).toISOString().slice(0, 10);
        return new Date(date).toISOString();
      }
    }
  });

  const result = ctx.buildOverviewFromTreatmentProgress_(
    [
      { patientId: '001', dateKey: '2025-02-01', time: '09:00' },
      { patientId: '002', dateKey: '2025-02-01', time: '10:00' },
      { patientId: '003', dateKey: '2025-01-31', time: '11:00' }
    ],
    new Date('2025-02-01T00:00:00Z'),
    'Asia/Tokyo'
  );

  assert.deepStrictEqual(JSON.parse(JSON.stringify(result)), {
    todayCount: 2,
    latestDayCount: 2
  });
}

function testVisitSummaryWhenNoDataReturnsZeroCounts() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        if (fmt === 'yyyy-MM-dd') return new Date(date).toISOString().slice(0, 10);
        return new Date(date).toISOString();
      }
    }
  });

  const result = ctx.buildOverviewFromTreatmentProgress_(
    [],
    new Date('2025-02-01T00:00:00Z'),
    'Asia/Tokyo'
  );

  assert.deepStrictEqual(JSON.parse(JSON.stringify(result)), {
    todayCount: 0,
    latestDayCount: 0
  });
}


function testInvoiceUnconfirmedUsesPositiveConfirmationEvidence() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const iso = new Date(date).toISOString();
        if (fmt === 'yyyy-MM') return iso.slice(0, 7);
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    }
  });

  const result = ctx.buildOverviewFromInvoiceUnconfirmed_(
    {},
    [
      { patientId: '001', dateKey: '2025-01-10', searchText: '前月施術あり' },
      { patientId: '001', dateKey: '2025-02-05', searchText: '請求書・領収書を受け渡し済み（家族へ）' },
      { patientId: '002', dateKey: '2025-01-08', searchText: '前月施術あり' },
      { patientId: '002', dateKey: '2025-02-22', searchText: '請求書・領収書を受け渡し済み' },
      { patientId: '003', dateKey: '2025-02-01', searchText: '当月のみ' }
    ],
    { notes: {} },
    { patientIds: new Set(['001', '002', '003']), applyFilter: true },
    { '001': '患者A', '002': '患者B', '003': '患者C' },
    new Date('2025-02-10T00:00:00Z'),
    'Asia/Tokyo'
  );

  assert.strictEqual(result.items.length, 1, '前月施術があり証跡がない患者のみ未対応になる');
  assert.strictEqual(result.items[0].patientId, '002');
}

function testInvoiceUnconfirmedIgnoresDisplayTargetFilter() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const iso = new Date(date).toISOString();
        if (fmt === 'yyyy-MM') return iso.slice(0, 7);
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    }
  });

  const result = ctx.getDashboardData({
    user: 'user@example.com',
    now: new Date('2025-02-10T00:00:00Z'),
    patientInfo: { patients: { '001': { name: '患者A' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { patientId: '001', timestamp: new Date('2025-01-10T00:00:00Z'), dateKey: '2025-01-10', createdByEmail: 'user@example.com', staffKeys: { email: 'user@example.com', name: '', staffId: '' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(result.overview.invoiceUnconfirmed.items.length, 1, 'displayTarget が空でも①請求の対象を保持する');
  assert.strictEqual(result.overview.invoiceUnconfirmed.items[0].patientId, '001');
}


function testInvoiceUnconfirmedShouldDetectPatientWithOnlyPreviousMonthTreatment() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const iso = new Date(date).toISOString();
        if (fmt === 'yyyy-MM') return iso.slice(0, 7);
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    }
  });

  const result = ctx.getDashboardData({
    user: 'user@example.com',
    now: new Date('2025-02-10T00:00:00Z'),
    patientInfo: { patients: { '001': { name: '患者A' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        {
          patientId: '001',
          timestamp: new Date('2025-01-10T00:00:00Z'),
          dateKey: '2025-01-10',
          searchText: '前月のみ施術',
          createdByEmail: 'user@example.com',
          staffKeys: { email: 'user@example.com', name: '', staffId: '' }
        }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(result.overview.invoiceUnconfirmed.items.length, 1, '前月のみ施術ログがある患者を未確認として検出する');
  assert.strictEqual(result.overview.invoiceUnconfirmed.items[0].patientId, '001');
}

function testInvoiceUnconfirmedExcludesMedicalAssistancePatient() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const iso = new Date(date).toISOString();
        if (fmt === 'yyyy-MM') return iso.slice(0, 7);
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    }
  });

  const result = ctx.getDashboardData({
    user: 'user@example.com',
    now: new Date('2025-02-10T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '患者A', AS: true }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        {
          patientId: '001',
          timestamp: new Date('2025-01-10T00:00:00Z'),
          dateKey: '2025-01-10',
          searchText: '前月のみ施術',
          createdByEmail: 'user@example.com',
          staffKeys: { email: 'user@example.com', name: '', staffId: '' }
        }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] },
    unpaidAlerts: { alerts: [], warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(result.overview.invoiceUnconfirmed.items.length, 0, '医療助成患者は請求未確認対象から除外する');
}


function testVisibleScopeForAdminShowsAllPatients() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const iso = new Date(date).toISOString();
        if (fmt === 'yyyy-MM') return iso.slice(0, 7);
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    },
    getTasks: opts => ({
      tasks: ['001', '002']
        .filter(pid => !opts.visiblePatientIds || opts.visiblePatientIds.has(pid))
        .map(pid => ({ patientId: pid, type: 'consentWarning' })),
      warnings: []
    }),
    getTodayVisits: opts => ({
      visits: ['001', '002']
        .filter(pid => !opts.visiblePatientIds || opts.visiblePatientIds.has(pid))
        .map(pid => ({ patientId: pid, dateKey: '2025-02-10', time: '09:00' })),
      warnings: []
    }),
    loadUnpaidAlerts: opts => ({
      alerts: ['001', '002']
        .filter(pid => !opts.visiblePatientIds || opts.visiblePatientIds.has(pid))
        .map(pid => ({ patientId: pid, patientName: `患者${pid}` })),
      warnings: []
    })
  });

  const result = ctx.getDashboardData({
    user: { email: 'staff@example.com', role: 'admin' },
    now: new Date('2025-02-13T00:00:00Z'),
    patientInfo: { patients: { '001': { name: '患者A' }, '002': { name: '患者B' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { patientId: '001', timestamp: new Date('2025-01-10T00:00:00Z'), dateKey: '2025-01-10', createdByEmail: 'staff@example.com', staffKeys: { email: 'staff@example.com', name: '', staffId: '' } },
        { patientId: '002', timestamp: new Date('2025-01-11T00:00:00Z'), dateKey: '2025-01-11', createdByEmail: 'staff@example.com', staffKeys: { email: 'staff@example.com', name: '', staffId: '' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] }
  });

  assert.strictEqual(result.patients.length, 2, '管理者は全患者表示する');
  assert.strictEqual(result.tasks.length, 0, '上段3ブロックは tasks 非依存とする');
  assert.strictEqual(result.todayVisits.length, 2, '管理者は訪問全件表示する');
  assert.strictEqual(result.unpaidAlerts.length, 2, '管理者は未回収アラート全件表示する');
  assert.strictEqual(result.overview.invoiceUnconfirmed.items.length, 2, '管理者は請求未確認を全患者分表示する');
}

function testVisibleScopeForStaffWithin50DaysShowsMatchedPatientsOnly() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        const iso = new Date(date).toISOString();
        if (fmt === 'yyyy-MM') return iso.slice(0, 7);
        if (fmt === 'yyyy-MM-dd') return iso.slice(0, 10);
        return iso;
      }
    },
    getTasks: opts => ({
      tasks: ['001', '002']
        .filter(pid => !opts.visiblePatientIds || opts.visiblePatientIds.has(pid))
        .map(pid => ({ patientId: pid, type: 'consentWarning' })),
      warnings: []
    }),
    getTodayVisits: opts => ({
      visits: ['001', '002']
        .filter(pid => !opts.visiblePatientIds || opts.visiblePatientIds.has(pid))
        .map(pid => ({ patientId: pid, dateKey: '2025-02-10', time: '09:00' })),
      warnings: []
    }),
    loadUnpaidAlerts: opts => ({
      alerts: ['001', '002']
        .filter(pid => !opts.visiblePatientIds || opts.visiblePatientIds.has(pid))
        .map(pid => ({ patientId: pid, patientName: `患者${pid}` })),
      warnings: []
    })
  });

  const result = ctx.getDashboardData({
    user: { email: 'staff@example.com', role: 'staff' },
    now: new Date('2025-02-13T00:00:00Z'),
    patientInfo: { patients: { '001': { name: '患者A' }, '002': { name: '患者B' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { patientId: '001', timestamp: new Date('2025-01-10T00:00:00Z'), dateKey: '2025-01-10', createdByEmail: 'staff@example.com', staffKeys: { email: 'staff@example.com', name: '', staffId: '' } },
        { patientId: '002', timestamp: new Date('2025-01-10T00:00:00Z'), dateKey: '2025-01-10', createdByEmail: 'other@example.com', staffKeys: { email: 'other@example.com', name: '', staffId: '' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] }
  });

  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.patients.map(p => p.patientId))), ['001'], 'スタッフは50日以内に記録した患者のみ表示する');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.tasks.map(t => t.patientId))), []);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.todayVisits.map(v => v.patientId))), ['001']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.unpaidAlerts.map(a => a.patientId))), ['001']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.overview.invoiceUnconfirmed.items.map(item => item.patientId))), ['001'], 'スタッフは請求未確認も担当患者のみ表示する');
}

function testVisibleScopeForStaffOnlyOlderThan50DaysShowsNoPatients() {
  const ctx = createContext({
    getTasks: opts => ({ tasks: opts.visiblePatientIds && opts.visiblePatientIds.size ? [{ patientId: '001' }] : [], warnings: [] }),
    getTodayVisits: opts => ({ visits: opts.visiblePatientIds && opts.visiblePatientIds.size ? [{ patientId: '001', dateKey: '2025-02-10', time: '09:00' }] : [], warnings: [] }),
    loadUnpaidAlerts: opts => ({ alerts: opts.visiblePatientIds && opts.visiblePatientIds.size ? [{ patientId: '001', patientName: '患者A' }] : [], warnings: [] })
  });

  const result = ctx.getDashboardData({
    user: { email: 'staff@example.com', role: 'staff' },
    now: new Date('2025-02-13T00:00:00Z'),
    patientInfo: { patients: { '001': { name: '患者A' } }, warnings: [] },
    notes: { notes: {}, warnings: [] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: {
      logs: [
        { patientId: '001', timestamp: new Date('2024-11-20T00:00:00Z'), createdByEmail: 'staff@example.com', staffKeys: { email: 'staff@example.com', name: '', staffId: '' } }
      ],
      warnings: []
    },
    responsible: { responsible: {}, warnings: [] }
  });

  assert.strictEqual(result.patients.length, 0, '50日より前のみなら患者表示0件');
  assert.strictEqual(result.tasks.length, 0);
  assert.strictEqual(result.todayVisits.length, 0);
  assert.strictEqual(result.unpaidAlerts.length, 0);
}

function testErrorIsCapturedInMeta() {
  const ctx = createContext({
    loadPatientInfo: () => { throw new Error('boom'); }
  });

  const result = ctx.getDashboardData();
  assert.strictEqual(result.tasks.length, 0);
  assert.strictEqual(result.todayVisits.length, 0);
  assert.strictEqual(result.patients.length, 0);
  assert.ok(result.meta.error && result.meta.error.indexOf('boom') >= 0);
}


function testSpreadsheetIsOpenedOnceAndPerfCheckIsLogged() {
  const logs = [];
  let openCount = 0;
  const workbook = { getSheetByName: () => null };
  const ctx = createContext({
    Logger: { log: message => logs.push(String(message)) },
    loadPatientInfo: opts => {
      assert.strictEqual(opts.dashboardSpreadsheet, workbook, 'loadPatientInfo に spreadsheet が伝搬する');
      return { patients: {}, nameToId: {}, warnings: [] };
    },
    loadNotes: opts => {
      assert.strictEqual(opts.dashboardSpreadsheet, workbook, 'loadNotes に spreadsheet が伝搬する');
      return { notes: {}, warnings: [] };
    },
    loadAIReports: opts => {
      assert.strictEqual(opts.dashboardSpreadsheet, workbook, 'loadAIReports に spreadsheet が伝搬する');
      return { reports: {}, warnings: [] };
    },
    loadInvoices: () => ({ invoices: {}, warnings: [] }),
    loadTreatmentLogs: opts => {
      assert.strictEqual(opts.dashboardSpreadsheet, workbook, 'loadTreatmentLogs に spreadsheet が伝搬する');
      return { logs: [], warnings: [], lastStaffByPatient: {} };
    },
    assignResponsibleStaff: opts => {
      assert.strictEqual(opts.dashboardSpreadsheet, workbook, 'assignResponsibleStaff に spreadsheet が伝搬する');
      return { responsible: {}, warnings: [] };
    },
    loadUnpaidAlerts: opts => {
      assert.strictEqual(opts.dashboardSpreadsheet, workbook, 'loadUnpaidAlerts に spreadsheet が伝搬する');
      return { alerts: [], warnings: [] };
    },
    getTasks: () => ({ tasks: [], warnings: [] }),
    getTodayVisits: () => ({ visits: [], warnings: [] })
  });

  ctx.dashboardGetSpreadsheet_ = () => { openCount += 1; return workbook; };

  ctx.getDashboardData({ user: 'user@example.com' });

  assert.strictEqual(openCount, 1, 'dashboardGetSpreadsheet_ は1回だけ呼ばれる');
  assert.ok(logs.includes('[perf-check] spreadsheetOpenCount=1'), 'perf-check ログが出力される');
}

function testWarningsAreDedupedAndSetupFlagged() {
  const warning = '患者情報シートが見つかりません';
  const ctx = createContext();
  const result = ctx.getDashboardData({
    patientInfo: { patients: {}, warnings: [warning], setupIncomplete: true },
    notes: { notes: {}, warnings: [warning] },
    aiReports: { reports: {}, warnings: [] },
    invoices: { invoices: {}, warnings: [] },
    treatmentLogs: { logs: [], warnings: [] },
    responsible: { responsible: {}, warnings: [] },
    tasksResult: { tasks: [], warnings: [] },
    visitsResult: { visits: [], warnings: [] }
  });

  assert.strictEqual(result.warnings.length, 1, '同一警告は一意になる');
  assert.strictEqual(result.warnings[0], warning);
  assert.strictEqual(result.meta.setupIncomplete, true, 'セットアップ未完了フラグが伝搬する');
}

(function run() {
  testAggregatesDashboardData();
  testPatientStatusTagsGeneration();
  testConsentOverviewMatchesPatientStatusTags();
  testEvaluateConsentStatusBoundaries();
  testConsentOver30DaysStillShowsReportTag();
  testEvaluateConsentStatusIsTimeIndependentForSameDate();
  testConsentAcquiredJudgmentHandlesFalseyStringsConsistently();
  testConsentDateParsingFormatsAndResolverPriority();
  testParseJapaneseEraDateAndResolveConsentExpiry();
  testConsentDateParseFailureCanBeDebugLogged();
  testStaffMatchingUsesEmailNameAndStaffIdWithLogs();
  testStaffConsentScopeMetricsAreLogged();
  testStaffConsentEligibilityEvaluatesOnlyVisibleOrMatchedPatients();
  testVisitSummaryWhenTodayIsZeroUsesLatestPastDayCount();
  testVisitSummaryWhenTodayHasTwoUsesTodayCountForBoth();
  testVisitSummaryWhenNoDataReturnsZeroCounts();
  testInvoiceUnconfirmedUsesPositiveConfirmationEvidence();
  testInvoiceUnconfirmedIgnoresDisplayTargetFilter();
  testInvoiceUnconfirmedShouldDetectPatientWithOnlyPreviousMonthTreatment();
  testInvoiceUnconfirmedExcludesMedicalAssistancePatient();
  testVisibleScopeForAdminShowsAllPatients();
  testVisibleScopeForStaffWithin50DaysShowsMatchedPatientsOnly();
  testVisibleScopeForStaffOnlyOlderThan50DaysShowsNoPatients();
  testErrorIsCapturedInMeta();
  testSpreadsheetIsOpenedOnceAndPerfCheckIsLogged();
  testWarningsAreDedupedAndSetupFlagged();
  console.log('dashboardGetDashboardData tests passed');
})();
