/**
 * ダッシュボードAPIの単体テストやデモ用に、スタブデータを生成する。
 * スプレッドシートやDriveがなくても getDashboardData() のレスポンスを確認できる。
 */
function getDashboardMockData(overrides) {
  const mock = buildDashboardMockData_(overrides);
  return typeof getDashboardData === 'function'
    ? getDashboardData(mock)
    : mock;
}

function buildDashboardMockData_(overrides) {
  const now = new Date('2025-12-15T09:30:00+09:00');
  const mockPatients = {
    P001: { patientId: 'P001', name: '施術 花子', consentExpiry: '2026/01/31' },
    P002: { patientId: 'P002', name: '山田 太郎', consentExpiry: '2025/12/25' }
  };

  const patientInfo = {
    patients: mockPatients,
    nameToId: Object.keys(mockPatients).reduce((map, pid) => {
      map[mockPatients[pid].name.replace(/\s+/g, '').toLowerCase()] = pid;
      return map;
    }, {}),
    warnings: []
  };

  const notes = {
    notes: {
      P001: {
        patientId: 'P001',
        preview: '昨日の施術後、肩の可動域が改善。継続観察。',
        when: '2025/12/14',
        unread: true,
        lastReadAt: '',
        authorEmail: 'writer@example.com'
      }
    },
    warnings: []
  };

  const aiReports = {
    reports: {
      P002: '2025/12/13'
    },
    warnings: []
  };

  const invoices = {
    invoices: {
      P001: 'https://example.com/invoice/P001',
      P002: null
    },
    warnings: []
  };

  const treatmentLogs = {
    logs: [
      {
        row: 2,
        patientId: 'P001',
        patientName: '施術 花子',
        createdByEmail: 'staff@example.com',
        timestamp: now,
        dateKey: '2025-12-15'
      },
      {
        row: 3,
        patientId: 'P002',
        patientName: '山田 太郎',
        createdByEmail: 'staff@example.com',
        timestamp: new Date('2025-12-14T10:00:00+09:00'),
        dateKey: '2025-12-14'
      }
    ],
    warnings: [],
    lastStaffByPatient: {
      P001: 'staff@example.com',
      P002: 'staff@example.com'
    }
  };

  const responsible = {
    responsible: {
      P001: '施術 太郎',
      P002: '施術 次郎'
    },
    warnings: []
  };

  const tasksResult = {
    tasks: [
      { type: 'consentWarning', severity: 'warning', patientId: 'P001', name: '施術 花子', detail: '同意期限が30日以内です' },
      { type: 'invoiceUnconfirmed', severity: 'warning', patientId: 'P002', name: '山田 太郎', detail: '請求書未確認' }
    ],
    warnings: []
  };

  const visitsResult = {
    visits: [
      { dateKey: '2025-12-15', time: '09:30', patientId: 'P001', patientName: '施術 花子', noteStatus: '◎' },
      { dateKey: '2025-12-14', time: '10:00', patientId: 'P002', patientName: '山田 太郎', noteStatus: '×' }
    ],
    warnings: []
  };

  return Object.assign({
    user: 'mock@example.com',
    now,
    patientInfo,
    treatmentLogs,
    notes,
    aiReports,
    invoices,
    responsible,
    tasksResult,
    visitsResult
  }, overrides || {});
}
