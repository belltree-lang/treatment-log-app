/**
 * ダッシュボードの主要データをまとめて取得し、JSON 形式で返す。
 * エラーが発生した場合は meta.error にメッセージを格納する。
 * @param {Object} [options]
 * @return {{tasks: Object[], todayVisits: Object[], patients: Object[], warnings: string[], meta: Object}}
 */
function getDashboardData(options) {
  const opts = options || {};
  let spreadsheetOpenCount = 0;
  const perfStartedAt = Date.now();
  const perfSteps = [];
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
  };
  const logPerf = message => {
    if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      Logger.log(message);
    } else if (typeof dashboardWarn_ === 'function') {
      dashboardWarn_(message);
    } else {
      logContext('perf', message);
    }
  };
  const measureStep = (step, runner) => {
    const start = Date.now();
    const result = runner();
    const duration = Date.now() - start;
    perfSteps.push({ step, duration });
    logPerf(`[perf] step=${step} duration=${duration}ms`);
    return result;
  };
  if (opts.mock && typeof buildDashboardMockData_ === 'function') {
    const mockOptions = buildDashboardMockData_(opts) || {};
    const normalized = Object.assign({}, mockOptions);
    normalized.mock = '';
    return getDashboardData(normalized);
  }
  const meta = {
    generatedAt: dashboardFormatDate_(new Date(), dashboardResolveTimeZone_(), "yyyy-MM-dd'T'HH:mm:ssXXX") || new Date().toISOString(),
    user: opts.user || dashboardResolveUser_(),
    setupIncomplete: false
  };
  const user = opts.user && typeof opts.user === 'object' ? opts.user : null;
  const userIdentity = user && user.email ? user.email : meta.user;
  meta.user = userIdentity || '';
  const normalizedUser = dashboardNormalizeEmail_(userIdentity || '');
  const isAdmin = Boolean(user && user.role === 'admin');
  logContext('getDashboardData:start', `user=${userIdentity || 'unknown'} normalizedUser=${normalizedUser || 'unknown'} isAdmin=${isAdmin ? 'true' : 'false'} mock=${opts.mock ? 'true' : 'false'}`);

  try {
    const dashboardSpreadsheet = measureStep('dashboardGetSpreadsheet', () => {
      if (opts.dashboardSpreadsheet) return opts.dashboardSpreadsheet;
      if (typeof dashboardGetSpreadsheet_ !== 'function') return null;
      spreadsheetOpenCount += 1;
      return dashboardGetSpreadsheet_();
    });
    const patientInfo = measureStep('loadPatientInfo', () => (opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo({ dashboardSpreadsheet }) : { patients: {}, nameToId: {}, warnings: [] })));
    logContext('getDashboardData:loadPatientInfo', `patients=${Object.keys(patientInfo && patientInfo.patients ? patientInfo.patients : {}).length} warnings=${(patientInfo && patientInfo.warnings ? patientInfo.warnings.length : 0)} setupIncomplete=${!!(patientInfo && patientInfo.setupIncomplete)}`);
    const notes = measureStep('loadNotes', () => (opts.notes || (typeof loadNotes === 'function' ? loadNotes({ email: userIdentity, dashboardSpreadsheet }) : { notes: {}, warnings: [] })));
    logContext('getDashboardData:loadNotes', `notes=${Object.keys(notes && notes.notes ? notes.notes : {}).length} warnings=${(notes && notes.warnings ? notes.warnings.length : 0)} setupIncomplete=${!!(notes && notes.setupIncomplete)}`);
    const aiReports = measureStep('loadAIReports', () => (opts.aiReports || (typeof loadAIReports === 'function' ? loadAIReports({ dashboardSpreadsheet }) : { reports: {}, warnings: [] })));
    logContext('getDashboardData:loadAIReports', `reports=${Object.keys(aiReports && aiReports.reports ? aiReports.reports : {}).length} warnings=${(aiReports && aiReports.warnings ? aiReports.warnings.length : 0)} setupIncomplete=${!!(aiReports && aiReports.setupIncomplete)}`);
    const invoices = measureStep('loadInvoices', () => (opts.invoices || (typeof loadInvoices === 'function'
      ? loadInvoices({ patientInfo, now: opts.now, cache: opts.cache, includePreviousMonth: true })
      : { invoices: {}, warnings: [] })));
    logContext('getDashboardData:loadInvoices', `invoices=${Object.keys(invoices && invoices.invoices ? invoices.invoices : {}).length} warnings=${(invoices && invoices.warnings ? invoices.warnings.length : 0)} setupIncomplete=${!!(invoices && invoices.setupIncomplete)}`);
    const treatmentLogs = measureStep('loadTreatmentLogs', () => (opts.treatmentLogs || (typeof loadTreatmentLogs === 'function'
      ? loadTreatmentLogs({ patientInfo, now: opts.now, cache: opts.cache, dashboardSpreadsheet })
      : { logs: [], warnings: [] })));
    const totalTreatmentLogs = treatmentLogs && treatmentLogs.logs ? treatmentLogs.logs.length : 0;
    logContext('getDashboardData:loadTreatmentLogs', `logs=${totalTreatmentLogs} warnings=${(treatmentLogs && treatmentLogs.warnings ? treatmentLogs.warnings.length : 0)} setupIncomplete=${!!(treatmentLogs && treatmentLogs.setupIncomplete)}`);
    const normalizedUserName = dashboardNormalizeStaffKey_(userIdentity || '');
    const normalizedUserId = dashboardNormalizeStaffKey_(userIdentity || '');
    const staffMatchResult = measureStep('staffMatch処理', () => {
      const staffMatchedLogs = [];
      const matchStats = { email: 0, name: 0, staffId: 0, none: 0 };
      const matchSamples = [];
      (treatmentLogs && Array.isArray(treatmentLogs.logs) ? treatmentLogs.logs : []).forEach(entry => {
        const staffKeys = entry && entry.staffKeys ? entry.staffKeys : {};
        const emailKey = dashboardNormalizeStaffKey_(staffKeys.email || (entry && entry.createdByEmail ? entry.createdByEmail : ''));
        const nameKey = dashboardNormalizeStaffKey_(staffKeys.name || (entry && entry.staffName ? entry.staffName : ''));
        const staffIdKey = dashboardNormalizeStaffKey_(staffKeys.staffId || (entry && entry.staffId ? entry.staffId : ''));

        let strategy = 'none';
        if (normalizedUser && emailKey && emailKey === normalizedUser) {
          strategy = 'email';
        } else if (normalizedUserName && nameKey && nameKey === normalizedUserName) {
          strategy = 'name';
        } else if (normalizedUserId && staffIdKey && staffIdKey === normalizedUserId) {
          strategy = 'staffId';
        }

        if (strategy !== 'none') {
          const matchedEntry = Object.assign({}, entry, { staffMatchStrategy: strategy });
          staffMatchedLogs.push(matchedEntry);
        }
        matchStats[strategy] += 1;
        if (matchSamples.length < 20 && strategy !== 'none') {
          matchSamples.push({
            row: entry && entry.row ? entry.row : null,
            staffMatchStrategy: strategy,
            emailKey,
            nameKey,
            staffIdKey
          });
        }
      });
      const matchedPatientIds = new Set();
      staffMatchedLogs.forEach(entry => {
        const pid = dashboardNormalizePatientId_(entry && entry.patientId);
        if (pid) matchedPatientIds.add(pid);
      });
      return { staffMatchedLogs, matchStats, matchSamples, matchedPatientIds };
    });
    const staffMatchedLogs = staffMatchResult.staffMatchedLogs;
    const matchStats = staffMatchResult.matchStats;
    const matchSamples = staffMatchResult.matchSamples;
    const matchedPatientIds = staffMatchResult.matchedPatientIds;
    const patientMaster = patientInfo && patientInfo.patients ? patientInfo.patients : {};
    const matchedPatientIdsInMaster = Array.from(matchedPatientIds).filter(pid => Object.prototype.hasOwnProperty.call(patientMaster, pid));
    logContext('getDashboardData:staffMatchStrategy', JSON.stringify({
      normalizedUser: normalizedUser || 'unknown',
      normalizedUserName: normalizedUserName || 'unknown',
      normalizedUserId: normalizedUserId || 'unknown',
      matchStats,
      samples: matchSamples
    }));
    logContext('getDashboardData:staffMatchedLogs', String(staffMatchedLogs.length));
    logContext('getDashboardData:matchedPatientIds', JSON.stringify(Array.from(matchedPatientIds)));
    logContext(
      'getDashboardData:staffMatchSummary',
      `normalizedUser=${normalizedUser || 'unknown'} totalLogs=${totalTreatmentLogs} staffMatchedLogs=${staffMatchedLogs.length} matchedPatientIds=${matchedPatientIds.size} matchedPatientIdsInMaster=${matchedPatientIdsInMaster.length}`
    );
    const responsible = measureStep('assignResponsible', () => (opts.responsible || (typeof assignResponsibleStaff === 'function'
      ? assignResponsibleStaff({ patientInfo, treatmentLogs, now: opts.now, cache: opts.cache, dashboardSpreadsheet })
      : { responsible: {}, warnings: [] })));
    logContext('getDashboardData:assignResponsible', `responsible=${Object.keys(responsible && responsible.responsible ? responsible.responsible : {}).length} warnings=${(responsible && responsible.warnings ? responsible.warnings.length : 0)} setupIncomplete=${!!(responsible && responsible.setupIncomplete)}`);

    const now = dashboardCoerceDate_(opts.now) || new Date();
    const fiftyDaysAgo = new Date(now.getTime() - 50 * 24 * 60 * 60 * 1000);
    const responsiblePatientIds = new Set();
    staffMatchedLogs.forEach(entry => {
      if (!entry || !entry.timestamp) return;
      const timestamp = dashboardCoerceDate_(entry.timestamp);
      if (!timestamp || timestamp.getTime() < fiftyDaysAgo.getTime()) return;
      const patientId = dashboardNormalizePatientId_(entry.patientId);
      if (!patientId) return;
      responsiblePatientIds.add(patientId);
    });
    const visiblePatientIds = isAdmin ? null : responsiblePatientIds;

    const unpaidAlertsResult = measureStep('loadUnpaidAlerts', () => (opts.unpaidAlerts || (typeof loadUnpaidAlerts === 'function'
      ? loadUnpaidAlerts({ patientInfo, now: opts.now, cache: opts.cache, dashboardSpreadsheet, visiblePatientIds })
      : { alerts: [], warnings: [] })));
    logContext('getDashboardData:loadUnpaidAlerts', `alerts=${(unpaidAlertsResult && unpaidAlertsResult.alerts ? unpaidAlertsResult.alerts.length : 0)} warnings=${(unpaidAlertsResult && unpaidAlertsResult.warnings ? unpaidAlertsResult.warnings.length : 0)} setupIncomplete=${!!(unpaidAlertsResult && unpaidAlertsResult.setupIncomplete)}`);

    const tasksResult = measureStep('getTasks', () => (opts.tasksResult || (typeof getTasks === 'function' ? getTasks({
      patientInfo,
      notes,
      aiReports,
      invoiceConfirmations: opts.invoiceConfirmations,
      visiblePatientIds,
      now: opts.now
    }) : { tasks: [], warnings: [] })));
    logContext('getDashboardData:getTasks', `tasks=${(tasksResult && tasksResult.tasks ? tasksResult.tasks.length : 0)} warnings=${(tasksResult && tasksResult.warnings ? tasksResult.warnings.length : 0)} setupIncomplete=${!!(tasksResult && tasksResult.setupIncomplete)}`);

    const visitsResult = measureStep('getTodayVisits', () => (opts.visitsResult || (typeof getTodayVisits === 'function' ? getTodayVisits({
      treatmentLogs,
      patientInfo,
      notes,
      visiblePatientIds,
      now: opts.now
    }) : { visits: [], warnings: [] })));
    logContext('getDashboardData:getTodayVisits', `visits=${(visitsResult && visitsResult.visits ? visitsResult.visits.length : 0)} warnings=${(visitsResult && visitsResult.warnings ? visitsResult.warnings.length : 0)} setupIncomplete=${!!(visitsResult && visitsResult.setupIncomplete)}`);

    const patients = measureStep('buildPatients', () => buildDashboardPatients_(patientInfo, {
      notes,
      aiReports,
      invoices,
      responsible,
      treatmentLogs
    }, visiblePatientIds));
    logContext('getDashboardData:buildPatients', `patients=${patients.length}`);

    const warningState = collectDashboardWarnings_([
      patientInfo,
      notes,
      aiReports,
      invoices,
      treatmentLogs,
      responsible,
      unpaidAlertsResult,
      tasksResult,
      visitsResult
    ]);

    meta.setupIncomplete = warningState.setupIncomplete;
    logContext('getDashboardData:setupIncomplete', `result=${meta.setupIncomplete} warnings=${warningState.warnings.length}`);

    const overview = buildDashboardOverview_({
      tasks: tasksResult && tasksResult.tasks ? tasksResult.tasks : [],
      visits: visitsResult && visitsResult.visits ? visitsResult.visits : [],
      patients,
      patientInfo,
      treatmentLogs,
      invoices,
      notes,
      user: userIdentity,
      now: opts.now,
      allowedPatientIds: visiblePatientIds
    });
    const invoiceUnconfirmed = overview && overview.invoiceUnconfirmed ? overview.invoiceUnconfirmed : { count: 0, items: [] };
    logContext('getDashboardData:overviewInvoiceUnconfirmed', `count=${Number(invoiceUnconfirmed.count) || 0} items.length=${Array.isArray(invoiceUnconfirmed.items) ? invoiceUnconfirmed.items.length : 0}`);
    return {
      tasks: tasksResult && tasksResult.tasks ? tasksResult.tasks : [],
      todayVisits: visitsResult && visitsResult.visits ? visitsResult.visits : [],
      patients,
      unpaidAlerts: unpaidAlertsResult && unpaidAlertsResult.alerts ? unpaidAlertsResult.alerts : [],
      warnings: warningState.warnings,
      overview,
      meta
    };
  } catch (err) {
    meta.error = err && err.message ? err.message : String(err);
    logContext('getDashboardData:error', meta.error);
    return { tasks: [], todayVisits: [], patients: [], unpaidAlerts: [], warnings: [], overview: null, meta };
  } finally {
    const totalDuration = Date.now() - perfStartedAt;
    logPerf(`[perf] total=${totalDuration}ms`);

    const plannedSteps = [
      'dashboardGetSpreadsheet',
      'loadPatientInfo',
      'loadNotes',
      'loadAIReports',
      'loadInvoices',
      'loadTreatmentLogs',
      'staffMatch処理',
      'assignResponsible',
      'loadUnpaidAlerts',
      'getTasks',
      'getTodayVisits',
      'buildPatients'
    ];
    const perfMap = {};
    perfSteps.forEach(entry => {
      perfMap[entry.step] = entry.duration;
    });
    const summaryLines = plannedSteps.map(step => {
      const duration = Object.prototype.hasOwnProperty.call(perfMap, step) ? `${perfMap[step]}ms` : 'not_run';
      return `  - ${step}: ${duration}`;
    });
    const top3 = perfSteps
      .slice()
      .sort((a, b) => b.duration - a.duration)
      .slice(0, 3)
      .map((entry, index) => `  - ${index + 1}. ${entry.step}: ${entry.duration}ms`);
    const top3Lines = top3.length ? top3 : ['  - 計測データなし'];

    logPerf('■ 処理時間サマリー');
    summaryLines.forEach(line => logPerf(line));
    logPerf(`  - 合計処理時間: ${totalDuration}ms`);

    logPerf('■ 最重処理トップ3');
    top3Lines.forEach(line => logPerf(line));

    logPerf('■ 5秒以内にするための削減候補');
    logPerf('  - 何を: loadTreatmentLogs の対象行数と列数 / どう削るか: 直近期間で先に絞り込み、必要列のみを読み込んで全件走査を削減する。');
    logPerf('  - 何を: loadAIReports・loadInvoices の逐次読み込み / どう削るか: キャッシュ（更新時刻付き）を導入して差分時のみ再計算する。');
    logPerf('  - 何を: staffMatch処理の全ログループ照合 / どう削るか: 正規化済みキーでインデックスを構築して比較回数を減らす。');

    logPerf('■ 副作用リスク');
    logPerf('  - データ欠損: 期間・列の絞り込みを誤ると過去必要データが取り込まれず、患者情報や請求情報が欠ける。');
    logPerf('  - 表示整合性: キャッシュや差分更新が古いと、タスク件数や未払いアラート件数が画面と実データで乖離する。');
    logPerf('  - 権限影響: ユーザー別フィルタを前段キャッシュへ寄せると、権限境界を跨いだデータ混入リスクが上がる。');
    logPerf(`[perf-check] spreadsheetOpenCount=${spreadsheetOpenCount}`);
  }
}

function buildDashboardPatients_(patientInfo, sources, allowedPatientIds) {
  const patients = [];
  const basePatients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const notes = sources && sources.notes && sources.notes.notes ? sources.notes.notes : {};
  const aiReports = sources && sources.aiReports && sources.aiReports.reports ? sources.aiReports.reports : {};
  const invoices = sources && sources.invoices && sources.invoices.invoices ? sources.invoices.invoices : {};
  const responsible = sources && sources.responsible && sources.responsible.responsible ? sources.responsible.responsible : {};

  const seen = new Set();
  const addPatient = (pid, payload) => {
    const patientId = dashboardNormalizePatientId_(pid);
    if (!patientId || seen.has(patientId)) return;
    if (allowedPatientIds && !allowedPatientIds.has(patientId)) return;
    if (!Object.prototype.hasOwnProperty.call(basePatients, patientId)) return;
    seen.add(patientId);

    const base = payload || {};
    const entry = {
      patientId,
      name: base.name || base.patientName || '',
      consentExpiry: base.consentExpiry || (base.raw && (base.raw['同意期限'] || base.raw['同意有効期限'])) || '',
      responsible: Object.prototype.hasOwnProperty.call(responsible, patientId) ? responsible[patientId] : null,
      invoiceUrl: Object.prototype.hasOwnProperty.call(invoices, patientId) ? invoices[patientId] : null,
      aiReportAt: Object.prototype.hasOwnProperty.call(aiReports, patientId) ? aiReports[patientId] : null,
      note: normalizeDashboardNote_(notes[patientId], patientId)
    };
    patients.push(entry);
  };

  Object.keys(basePatients).forEach(pid => addPatient(pid, basePatients[pid]));

  return patients;
}

function normalizeDashboardNote_(note, patientId) {
  if (!note) return null;
  return {
    patientId: patientId || note.patientId || '',
    preview: note.preview || '',
    when: note.when || '',
    unread: !!note.unread,
    lastReadAt: note.lastReadAt || '',
    authorEmail: note.authorEmail || '',
    row: note.row || null
  };
}

function buildDashboardOverview_(params) {
  const payload = params || {};
  const tasks = Array.isArray(payload.tasks) ? payload.tasks : [];
  const patients = Array.isArray(payload.patients) ? payload.patients : [];
  const patientInfo = payload.patientInfo && payload.patientInfo.patients ? payload.patientInfo.patients : {};
  const invoices = payload.invoices || {};
  const treatmentLogs = payload.treatmentLogs && Array.isArray(payload.treatmentLogs.logs)
    ? payload.treatmentLogs.logs
    : [];
  const now = dashboardCoerceDate_(payload.now) || new Date();
  const tz = dashboardResolveTimeZone_();
  const user = payload.user || '';
  const allowedPatientIds = payload.allowedPatientIds && typeof payload.allowedPatientIds.has === 'function' ? payload.allowedPatientIds : null;
  const scope = { patientIds: allowedPatientIds, applyFilter: !!allowedPatientIds };
  const invoiceScope = { patientIds: null, applyFilter: false };

  const patientNameMap = {};
  patients.forEach(entry => {
    if (entry && entry.patientId) {
      patientNameMap[entry.patientId] = entry.name || entry.patientName || '';
    }
  });
  Object.keys(patientInfo || {}).forEach(pid => {
    if (!pid || patientNameMap[pid]) return;
    const info = patientInfo[pid] || {};
    patientNameMap[pid] = info.name || info.patientName || '';
  });

  const invoiceUnconfirmed = buildOverviewFromInvoiceUnconfirmed_(
    tasks,
    invoices,
    treatmentLogs,
    payload.notes,
    invoiceScope,
    patientNameMap,
    now,
    tz,
    patientInfo
  );
  const consentRelated = buildOverviewFromConsent_(tasks, patientInfo, scope, patientNameMap, now, tz);
  const visitSummary = buildOverviewFromTreatmentProgress_(treatmentLogs, user, now, tz);

  return {
    invoiceUnconfirmed,
    consentRelated,
    visitSummary
  };
}

function buildOverviewFromInvoiceUnconfirmed_(tasks, invoices, treatmentLogs, notes, scope, patientNameMap, now, tz, patientInfo) {
  const items = [];
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;
  const targetNow = dashboardCoerceDate_(now) || new Date();
  const previousMonthKey = dashboardFormatDate_(new Date(targetNow.getFullYear(), targetNow.getMonth() - 1, 1), tz, 'yyyy-MM');
  const currentMonthKey = dashboardFormatDate_(targetNow, tz, 'yyyy-MM');
  const confirmationPhrase = '請求書・領収書を受け渡し済み';
  const logBillingDebug = (message, details) => {
    const payload = details ? ` ${details}` : '';
    const line = `[billing-debug] ${message}${payload}`;
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_('billing-debug', `${message}${payload}`);
    } else if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      Logger.log(line);
    } else if (typeof dashboardWarn_ === 'function') {
      dashboardWarn_(line);
    }
  };

  const prevMonthPatientIds = new Set();
  (treatmentLogs || []).forEach(entry => {
    const pid = dashboardNormalizePatientId_(entry && entry.patientId);
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    const monthKey = resolveDashboardMonthKey_(entry, tz);
    if (monthKey !== previousMonthKey) return;
    prevMonthPatientIds.add(pid);
  });
  logBillingDebug(`prevMonthPatientIds count=${prevMonthPatientIds.size}`);

  const confirmedPatients = new Set();
  const inBillingWindowPatients = new Set();
  (treatmentLogs || []).forEach(entry => {
    const pid = dashboardNormalizePatientId_(entry && entry.patientId);
    if (!pid || !prevMonthPatientIds.has(pid)) return;
    if (!isDashboardInvoiceConfirmationInWindow_(entry, currentMonthKey, tz)) return;
    inBillingWindowPatients.add(pid);
    const searchable = buildDashboardInvoiceSearchText_(entry);
    if (searchable.indexOf(confirmationPhrase) >= 0) confirmedPatients.add(pid);
  });

  const noteEntries = notes && notes.notes ? notes.notes : {};
  Object.keys(noteEntries).forEach(pidRaw => {
    const pid = dashboardNormalizePatientId_(pidRaw);
    if (!pid || !prevMonthPatientIds.has(pid) || confirmedPatients.has(pid)) return;
    const note = noteEntries[pidRaw] || {};
    if (!isDashboardInvoiceConfirmationInWindow_(note, currentMonthKey, tz)) return;
    inBillingWindowPatients.add(pid);
    const searchable = buildDashboardInvoiceSearchText_(note);
    if (searchable.indexOf(confirmationPhrase) >= 0) confirmedPatients.add(pid);
  });

  logBillingDebug(`inBillingWindow count=${inBillingWindowPatients.size}`);
  logBillingDebug(`completedByText count=${confirmedPatients.size}`);

  prevMonthPatientIds.forEach(pid => {
    if (isDashboardMedicalAssistancePatient_(patientInfo, pid)) return;
    if (confirmedPatients.has(pid)) return;
    items.push({
      patientId: pid,
      name: patientNameMap[pid] || '',
      count: 1,
      subText: `受渡未確認（対象月: ${previousMonthKey}）`
    });
  });

  const pendingSample = items.slice(0, 10).map(item => item.patientId);
  logBillingDebug(`pendingPatients count=${items.length}`, `sample=${JSON.stringify(pendingSample)}`);

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  return { count: items.length, items };
}

function isDashboardMedicalAssistancePatient_(patientInfo, patientId) {
  if (!patientInfo || !patientId) return false;
  const info = patientInfo[patientId] || {};
  if (info.AS === true) return true;
  const raw = info.raw || {};
  if (raw.AS === true) return true;
  if (raw['医療助成'] === true) return true;
  return false;
}

function resolveDashboardMonthKey_(entry, tz) {
  const dateKey = entry && entry.dateKey ? String(entry.dateKey).trim() : '';
  if (dateKey && /^\d{4}-\d{2}-\d{2}$/.test(dateKey)) return dateKey.slice(0, 7);
  const ts = dashboardCoerceDate_(entry && entry.timestamp ? entry.timestamp : (entry && entry.when ? entry.when : null));
  if (!ts) return '';
  return dashboardFormatDate_(ts, tz, 'yyyy-MM');
}

function isDashboardInvoiceConfirmationInWindow_(entry, currentMonthKey, tz) {
  const dateKey = entry && entry.dateKey ? String(entry.dateKey).trim() : '';
  let dayKey = dateKey;
  if (!dayKey || !/^\d{4}-\d{2}-\d{2}$/.test(dayKey)) {
    const ts = dashboardCoerceDate_(entry && entry.timestamp ? entry.timestamp : (entry && entry.when ? entry.when : null));
    if (!ts) return false;
    dayKey = dashboardFormatDate_(ts, tz, 'yyyy-MM-dd');
  }
  if (dayKey.slice(0, 7) !== currentMonthKey) return false;
  const day = Number(dayKey.slice(8, 10));
  return day >= 1 && day <= 20;
}

function buildDashboardInvoiceSearchText_(entry) {
  if (!entry || typeof entry !== 'object') return '';
  const candidates = [
    entry.text,
    entry.note,
    entry.memo,
    entry.body,
    entry.content,
    entry.summary,
    entry.searchText
  ];
  return candidates
    .map(value => String(value == null ? '' : value).trim())
    .filter(Boolean)
    .join('\n');
}

function buildOverviewFromConsent_(tasks, patientInfo, scope, patientNameMap, now, tz) {
  const items = [];
  const issuesByPatient = {};
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;

  (tasks || []).forEach(task => {
    const pid = task && task.patientId ? String(task.patientId).trim() : '';
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    const type = task && task.type ? String(task.type) : '';
    if (type !== 'consentExpired' && type !== 'consentWarning') return;
    const detail = task.detail ? String(task.detail) : '';
    const label = type === 'consentExpired' ? '期限切れ' : '期限間近';
    const message = detail ? `${label}: ${detail}` : label;
    if (!issuesByPatient[pid]) issuesByPatient[pid] = [];
    issuesByPatient[pid].push(message);
  });

  Object.keys(patientInfo || {}).forEach(pid => {
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    const info = patientInfo[pid] || {};
    const consentHandout = resolvePatientRawValue_(info.raw, [
      '配布', '配布欄', '配布状況', '配布日', '配布（同意書）', '同意書受渡', '同意書受け渡し', '同意書受渡日'
    ]);
    const consentDate = resolvePatientRawValue_(info.raw, [
      '同意年月日', '同意日', '同意開始日', '同意開始'
    ]);
    const visitPlanDate = resolvePatientRawValue_(info.raw, [
      '通院予定日', '通院日', '来院日', '通院予定', '来院予定'
    ]);

    if (consentHandout && !consentDate) {
      if (!issuesByPatient[pid]) issuesByPatient[pid] = [];
      issuesByPatient[pid].push('同意未取得');
    }
    if (consentHandout && !visitPlanDate) {
      if (!issuesByPatient[pid]) issuesByPatient[pid] = [];
      issuesByPatient[pid].push('通院日未定');
    }
  });

  Object.keys(issuesByPatient).forEach(pid => {
    const reasons = issuesByPatient[pid] || [];
    if (!reasons.length) return;
    const info = patientInfo[pid] || {};
    const name = info.name || patientNameMap[pid] || '';
    const subText = reasons.join(' / ');
    items.push({
      patientId: pid,
      name,
      count: Math.max(1, reasons.length),
      subText
    });
  });

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  return { count: items.length, items };
}

function resolvePatientRawValue_(raw, candidates) {
  if (!raw) return '';
  const keys = Array.isArray(candidates) ? candidates : [];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (Object.prototype.hasOwnProperty.call(raw, key)) {
      const value = raw[key];
      if (value != null && String(value).trim()) return value;
    }
  }
  return '';
}

function buildOverviewFromTreatmentProgress_(treatmentLogs, user, now, tz) {
  const targetNow = dashboardCoerceDate_(now) || new Date();
  const todayKey = dashboardFormatDate_(targetNow, tz, 'yyyy-MM-dd');
  const normalizedUser = dashboardNormalizeEmail_(user || '');
  const filtered = (treatmentLogs || []).filter(entry => {
    if (!normalizedUser) return true;
    const entryEmail = dashboardNormalizeEmail_(entry && entry.createdByEmail ? entry.createdByEmail : '');
    return entryEmail && entryEmail === normalizedUser;
  });

  const countsByDate = {};
  filtered.forEach(entry => {
    const key = entry && entry.dateKey ? String(entry.dateKey).trim() : '';
    if (!key) return;
    countsByDate[key] = (countsByDate[key] || 0) + 1;
  });

  const todayCount = countsByDate[todayKey] || 0;
  const dateKeys = Object.keys(countsByDate).sort();
  const latestDateKey = dateKeys.length ? dateKeys[dateKeys.length - 1] : '';
  const recentOneDayCount = latestDateKey ? (countsByDate[latestDateKey] || 0) : 0;

  return { todayCount, recentOneDayCount };
}

function collectDashboardWarnings_(results) {
  const warnings = [];
  const seen = new Set();
  let setupIncomplete = false;

  (results || []).forEach(entry => {
    if (entry && entry.setupIncomplete) {
      setupIncomplete = true;
      if (typeof dashboardLogContext_ === 'function') {
        dashboardLogContext_('collectDashboardWarnings', 'setupIncomplete from entry');
      } else if (typeof dashboardWarn_ === 'function') {
        dashboardWarn_('[collectDashboardWarnings] setupIncomplete from entry');
      }
    }
    if (entry && Array.isArray(entry.warnings)) {
      entry.warnings.forEach(warning => {
        const normalized = normalizeDashboardWarning_(warning);
        if (!normalized || seen.has(normalized)) return;
        seen.add(normalized);
        warnings.push(normalized);
        if (isDashboardSetupIncompleteWarning_(normalized)) {
          setupIncomplete = true;
          if (typeof dashboardLogContext_ === 'function') {
            dashboardLogContext_('collectDashboardWarnings', `setupIncomplete warning=${normalized}`);
          } else if (typeof dashboardWarn_ === 'function') {
            dashboardWarn_(`[collectDashboardWarnings] setupIncomplete warning=${normalized}`);
          }
        }
      });
    }
  });

  return { warnings, setupIncomplete };
}

function normalizeDashboardWarning_(warning) {
  const normalized = String(warning == null ? '' : warning).trim();
  if (!normalized) return '';
  if (shouldSuppressDashboardWarning_(normalized)) return '';
  return normalized;
}

function isDashboardSetupIncompleteWarning_(warning) {
  const normalized = normalizeDashboardWarning_(warning);
  if (!normalized) return false;
  if (normalized.indexOf('シートが見つかりません') >= 0) return true;
  if (normalized.indexOf('スプレッドシートを取得できませんでした') >= 0) return true;
  if (normalized.indexOf('請求書フォルダが取得できませんでした') >= 0) return true;
  return false;
}

function shouldSuppressDashboardWarning_(warning) {
  const normalized = String(warning == null ? '' : warning);
  if (!normalized) return true;
  if (/(row[:：]\s*\d+)/i.test(normalized)) return true;
  if (normalized.indexOf('患者ID') >= 0 && (normalized.indexOf('空') >= 0 || normalized.indexOf('未入力') >= 0)) return true;
  if (normalized.indexOf('患者名をIDに紐付け') >= 0) return true;
  if (normalized.indexOf('氏名が未入力') >= 0) return true;
  return false;
}

function resolveDashboardDisplayTargets_(params) {
  const payload = params || {};
  const patientInfo = payload.patientInfo && payload.patientInfo.patients ? payload.patientInfo.patients : {};
  const treatmentLogs = payload.treatmentLogs && Array.isArray(payload.treatmentLogs.logs)
    ? payload.treatmentLogs.logs
    : [];
  const now = dashboardCoerceDate_(payload.now) || new Date();
  const tz = dashboardResolveTimeZone_();
  const monthStart = dashboardStartOfMonth_(tz, now);
  const todayKey = dashboardFormatDate_(now, tz, 'yyyy-MM-dd');
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const yesterdayKey = dashboardFormatDate_(yesterday, tz, 'yyyy-MM-dd');
  const allowed = new Set();
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
  };
  let scopedCount = 0;
  let monthlyCount = 0;
  let dailyCount = 0;
  let totalCount = 0;

  (treatmentLogs || []).forEach(entry => {
    if (!entry || !entry.timestamp) return;
    totalCount += 1;
    const ts = dashboardCoerceDate_(entry.timestamp);
    if (!ts) return;
    const dateKey = entry.dateKey || dashboardFormatDate_(ts, tz, 'yyyy-MM-dd');
    const inMonthlyScope = ts >= monthStart;
    const inDailyScope = dateKey === todayKey || dateKey === yesterdayKey;
    if (!inMonthlyScope && !inDailyScope) return;
    if (inMonthlyScope) monthlyCount += 1;
    if (inDailyScope) dailyCount += 1;
    scopedCount += 1;
    const pid = dashboardNormalizePatientId_(entry.patientId);
    if (!pid || !Object.prototype.hasOwnProperty.call(patientInfo, pid)) return;
    allowed.add(pid);
  });

  logContext(
    'resolveDashboardDisplayTargets:summary',
    `totalLogs=${totalCount} inMonthlyScope=${monthlyCount} inDailyScope=${dailyCount} scopedLogs=${scopedCount} allowedPatientIds=${allowed.size}`
  );
  return { patientIds: allowed };
}

function filterDashboardItemsByPatientId_(items, allowedPatientIds) {
  if (!Array.isArray(items)) return [];
  if (!allowedPatientIds || typeof allowedPatientIds.has !== 'function') return items.slice();
  return items.filter(item => {
    const pid = dashboardNormalizePatientId_(item && item.patientId);
    return pid && allowedPatientIds.has(pid);
  });
}

function dashboardResolveUser_() {
  if (typeof Session !== 'undefined' && Session && typeof Session.getActiveUser === 'function') {
    try {
      const email = Session.getActiveUser().getEmail();
      if (email) return String(email).trim();
    } catch (e) { /* ignore */ }
  }
  return '';
}
