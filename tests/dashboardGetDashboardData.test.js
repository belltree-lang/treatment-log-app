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
    patients: { '001': { name: '山田太郎', consentExpiry: '2025-01-31' } },
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
  const visitsResult = { visits: [{ patientId: '001', time: '10:00' }], warnings: ['visit'] };
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
  assert.deepStrictEqual(result.tasks, tasksResult.tasks);
  assert.deepStrictEqual(result.todayVisits, visitsResult.visits);
  assert.strictEqual(result.patients.length, 1);
  const normalizedPatient = JSON.parse(JSON.stringify(result.patients[0]));
  assert.deepStrictEqual(normalizedPatient, {
    patientId: '001',
    name: '山田太郎',
    consentExpiry: '2025-01-31',
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
      { type: 'consent-expiry', level: 'danger', label: '期限超過', priority: 3 },
      { type: 'ai-report-at', level: 'info', label: 'AI報告日時', priority: 1 },
      { type: 'responsible', level: 'info', label: '担当者', priority: 1 },
      { type: 'unread-note', level: 'info', label: '未読ヒント', priority: 1 }
    ]
  });
  assert.strictEqual(result.unpaidAlerts.length, 1, '未回収アラートが伝搬する');
  assert.strictEqual(result.unpaidAlerts[0].patientId, '001');
  const warnings = JSON.parse(JSON.stringify(result.warnings)).sort();
  assert.deepStrictEqual(warnings, ['a1', 'i1', 'n1', 'p1', 'r1', 't1', 'task', 'u1', 'visit'].sort());
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.overview.patientStatusSummary)), {
    consentExpiredCount: 1,
    reportDelayedCount: 0
  }, '優先度集計をoverviewへ含める');
}



function testPatientStatusTagsGeneration() {
  const ctx = createContext();
  const result = ctx.getDashboardData({
    user: { email: 'user@example.com', role: 'admin' },
    now: new Date('2025-02-01T00:00:00Z'),
    patientInfo: {
      patients: {
        '001': { name: '期限超過', consentExpiry: '2025-01-20', raw: { '同意書受渡': '済' } },
        '002': { name: '期限間近', consentExpiry: '2025-02-20', raw: { '通院日未定': 'はい' } },
        '003': { name: '取得確認済', consentExpiry: '2025-03-20', raw: { '同意書取得確認': '済', '同意書受渡': '済', '通院日未定': 'はい' } },
        '004': { name: '報告書遅延', consentExpiry: '', raw: {} },
        '005': { name: '報告書なし', consentExpiry: '', raw: {} }
      },
      warnings: []
    },
    notes: { notes: {}, warnings: [] },
    aiReports: {
      reports: {
        '001': '2025-01-30',
        '002': '2025-01-30',
        '003': '2025-01-30',
        '004': '2024-07-01'
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
    { type: 'consent-expiry', level: 'danger', label: '期限超過', priority: 3 },
    { type: 'ai-report-at', level: 'info', label: 'AI報告日時', priority: 1 },
    { type: 'consent', level: 'warning', label: '同意書受渡', priority: 1 }
  ], '期限超過と同意書受渡タグを付与する');

  assert.deepStrictEqual(patientsById['002'], [
    { type: 'consent-expiry', level: 'warning', label: '残19日', priority: 2 },
    { type: 'ai-report-at', level: 'info', label: 'AI報告日時', priority: 1 },
    { type: 'consent', level: 'warning', label: '通院日未定', priority: 1 }
  ], '期限間近と通院日未定タグを付与する');

  assert.deepStrictEqual(patientsById['003'], [
    { type: 'ai-report-at', level: 'info', label: 'AI報告日時', priority: 1 }
  ], '同意書取得確認がある場合は同意タグを付与せず、情報タグのみ付与する');

  assert.deepStrictEqual(patientsById['004'], [
    { type: 'report', level: 'warning', label: '報告書遅延', priority: 3 },
    { type: 'ai-report-at', level: 'info', label: 'AI報告日時', priority: 1 }
  ], '180日以上前の報告書は遅延タグを付与する');

  assert.deepStrictEqual(patientsById['005'], [
    { type: 'report', level: 'warning', label: '報告書未発行', priority: 2 }
  ], '報告書がない場合は未発行タグを付与する');
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


function testVisitSummaryUsesCountsForTodayAndRecentOneDay() {
  const ctx = createContext({
    Utilities: {
      formatDate: (date, _tz, fmt) => {
        if (fmt === 'yyyy-MM-dd') return new Date(date).toISOString().slice(0, 10);
        return new Date(date).toISOString();
      }
    }
  });

  const logs = [
    { dateKey: '2025-02-01', createdByEmail: 'user@example.com' },
    { dateKey: '2025-01-31', createdByEmail: 'user@example.com' },
    { dateKey: '2025-01-31', createdByEmail: 'user@example.com' },
    { dateKey: '2025-01-30', createdByEmail: 'other@example.com' }
  ];

  const resultToday = ctx.buildOverviewFromTreatmentProgress_(
    logs,
    'user@example.com',
    new Date('2025-02-01T00:00:00Z'),
    'Asia/Tokyo'
  );
  assert.strictEqual(resultToday.todayCount, 1, '今日の件数を返す');
  assert.strictEqual(resultToday.recentOneDayCount, 1, '直近1日施術は最新日の件数を返す');

  const resultNoToday = ctx.buildOverviewFromTreatmentProgress_(
    logs,
    'user@example.com',
    new Date('2025-02-02T00:00:00Z'),
    'Asia/Tokyo'
  );
  assert.strictEqual(resultNoToday.todayCount, 0, '今日が0件の場合は0件を返す');
  assert.strictEqual(resultNoToday.recentOneDayCount, 1, '今日0件の場合は過去で最も新しい施術日の件数を返す');
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
    [],
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

  assert.strictEqual(result.count, 1, '前月施術があり証跡がない患者のみ未対応になる');
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

  assert.strictEqual(result.overview.invoiceUnconfirmed.count, 1, 'displayTarget が空でも①請求の対象を保持する');
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

  assert.strictEqual(result.overview.invoiceUnconfirmed.count, 1, '前月のみ施術ログがある患者を未確認として検出する');
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

  assert.strictEqual(result.overview.invoiceUnconfirmed.count, 0, '医療助成患者は請求未確認対象から除外する');
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
  assert.strictEqual(result.tasks.length, 2, '管理者はタスク全件表示する');
  assert.strictEqual(result.todayVisits.length, 2, '管理者は訪問全件表示する');
  assert.strictEqual(result.unpaidAlerts.length, 2, '管理者は未回収アラート全件表示する');
  assert.strictEqual(result.overview.invoiceUnconfirmed.count, 2, '管理者は請求未確認を全患者分表示する');
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
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.tasks.map(t => t.patientId))), ['001']);
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
  testStaffMatchingUsesEmailNameAndStaffIdWithLogs();
  testVisitSummaryUsesCountsForTodayAndRecentOneDay();
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
