/**
 * ダッシュボードの主要データをまとめて取得し、JSON 形式で返す。
 * エラーが発生した場合は meta.error にメッセージを格納する。
 * @param {Object} [options]
 * @return {{tasks: Object[], todayVisits: Object[], patients: Object[], warnings: string[], meta: Object}}
 */
function getDashboardData(options) {
  const opts = options || {};
  const logContext = (label, details) => {
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_(label, details);
    } else if (typeof dashboardWarn_ === 'function') {
      const payload = details ? ` ${details}` : '';
      dashboardWarn_(`[${label}]${payload}`);
    }
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
  logContext('getDashboardData:start', `user=${meta.user || 'unknown'} mock=${opts.mock ? 'true' : 'false'}`);

  try {
    const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : { patients: {}, nameToId: {}, warnings: [] });
    logContext('getDashboardData:loadPatientInfo', `patients=${Object.keys(patientInfo && patientInfo.patients ? patientInfo.patients : {}).length} warnings=${(patientInfo && patientInfo.warnings ? patientInfo.warnings.length : 0)} setupIncomplete=${!!(patientInfo && patientInfo.setupIncomplete)}`);
    const notes = opts.notes || (typeof loadNotes === 'function' ? loadNotes({ email: meta.user }) : { notes: {}, warnings: [] });
    logContext('getDashboardData:loadNotes', `notes=${Object.keys(notes && notes.notes ? notes.notes : {}).length} warnings=${(notes && notes.warnings ? notes.warnings.length : 0)} setupIncomplete=${!!(notes && notes.setupIncomplete)}`);
    const aiReports = opts.aiReports || (typeof loadAIReports === 'function' ? loadAIReports() : { reports: {}, warnings: [] });
    logContext('getDashboardData:loadAIReports', `reports=${Object.keys(aiReports && aiReports.reports ? aiReports.reports : {}).length} warnings=${(aiReports && aiReports.warnings ? aiReports.warnings.length : 0)} setupIncomplete=${!!(aiReports && aiReports.setupIncomplete)}`);
    const invoices = opts.invoices || (typeof loadInvoices === 'function'
      ? loadInvoices({ patientInfo, now: opts.now })
      : { invoices: {}, warnings: [] });
    logContext('getDashboardData:loadInvoices', `invoices=${Object.keys(invoices && invoices.invoices ? invoices.invoices : {}).length} warnings=${(invoices && invoices.warnings ? invoices.warnings.length : 0)} setupIncomplete=${!!(invoices && invoices.setupIncomplete)}`);
    const treatmentLogs = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function'
      ? loadTreatmentLogs({ patientInfo, now: opts.now })
      : { logs: [], warnings: [] });
    logContext('getDashboardData:loadTreatmentLogs', `logs=${(treatmentLogs && treatmentLogs.logs ? treatmentLogs.logs.length : 0)} warnings=${(treatmentLogs && treatmentLogs.warnings ? treatmentLogs.warnings.length : 0)} setupIncomplete=${!!(treatmentLogs && treatmentLogs.setupIncomplete)}`);
    const responsible = opts.responsible || (typeof assignResponsibleStaff === 'function'
      ? assignResponsibleStaff({ patientInfo, treatmentLogs, now: opts.now })
      : { responsible: {}, warnings: [] });
    logContext('getDashboardData:assignResponsible', `responsible=${Object.keys(responsible && responsible.responsible ? responsible.responsible : {}).length} warnings=${(responsible && responsible.warnings ? responsible.warnings.length : 0)} setupIncomplete=${!!(responsible && responsible.setupIncomplete)}`);
    const unpaidAlertsResult = opts.unpaidAlerts || (typeof loadUnpaidAlerts === 'function'
      ? loadUnpaidAlerts({ patientInfo, now: opts.now })
      : { alerts: [], warnings: [] });
    logContext('getDashboardData:loadUnpaidAlerts', `alerts=${(unpaidAlertsResult && unpaidAlertsResult.alerts ? unpaidAlertsResult.alerts.length : 0)} warnings=${(unpaidAlertsResult && unpaidAlertsResult.warnings ? unpaidAlertsResult.warnings.length : 0)} setupIncomplete=${!!(unpaidAlertsResult && unpaidAlertsResult.setupIncomplete)}`);

    const tasksResult = opts.tasksResult || (typeof getTasks === 'function' ? getTasks({
      patientInfo,
      notes,
      aiReports,
      invoiceConfirmations: opts.invoiceConfirmations,
      now: opts.now
    }) : { tasks: [], warnings: [] });
    logContext('getDashboardData:getTasks', `tasks=${(tasksResult && tasksResult.tasks ? tasksResult.tasks.length : 0)} warnings=${(tasksResult && tasksResult.warnings ? tasksResult.warnings.length : 0)} setupIncomplete=${!!(tasksResult && tasksResult.setupIncomplete)}`);

    const visitsResult = opts.visitsResult || (typeof getTodayVisits === 'function' ? getTodayVisits({
      treatmentLogs,
      notes,
      now: opts.now
    }) : { visits: [], warnings: [] });
    logContext('getDashboardData:getTodayVisits', `visits=${(visitsResult && visitsResult.visits ? visitsResult.visits.length : 0)} warnings=${(visitsResult && visitsResult.warnings ? visitsResult.warnings.length : 0)} setupIncomplete=${!!(visitsResult && visitsResult.setupIncomplete)}`);

    const displayTargets = resolveDashboardDisplayTargets_({
      patientInfo,
      treatmentLogs,
      now: opts.now
    });
    const filteredTasks = filterDashboardItemsByPatientId_(tasksResult && tasksResult.tasks ? tasksResult.tasks : [], displayTargets.patientIds);
    const filteredVisits = filterDashboardItemsByPatientId_(visitsResult && visitsResult.visits ? visitsResult.visits : [], displayTargets.patientIds);
    const filteredAlerts = filterDashboardItemsByPatientId_(unpaidAlertsResult && unpaidAlertsResult.alerts ? unpaidAlertsResult.alerts : [], displayTargets.patientIds);
    const filteredTreatmentLogs = filterDashboardItemsByPatientId_(treatmentLogs && Array.isArray(treatmentLogs.logs) ? treatmentLogs.logs : [], displayTargets.patientIds);

    const patients = buildDashboardPatients_(patientInfo, {
      notes,
      aiReports,
      invoices,
      responsible,
      treatmentLogs
    }, displayTargets.patientIds);
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

    return {
      tasks: filteredTasks,
      todayVisits: filteredVisits,
      patients,
      unpaidAlerts: filteredAlerts,
      warnings: warningState.warnings,
      overview: buildDashboardOverview_({
        tasks: filteredTasks,
        visits: filteredVisits,
        patients,
        patientInfo,
        treatmentLogs: { logs: filteredTreatmentLogs },
        user: meta.user,
        now: opts.now
      }),
      meta
    };
  } catch (err) {
    meta.error = err && err.message ? err.message : String(err);
    logContext('getDashboardData:error', meta.error);
    return { tasks: [], todayVisits: [], patients: [], unpaidAlerts: [], warnings: [], overview: null, meta };
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
  const visits = Array.isArray(payload.visits) ? payload.visits : [];
  const patients = Array.isArray(payload.patients) ? payload.patients : [];
  const patientInfo = payload.patientInfo && payload.patientInfo.patients ? payload.patientInfo.patients : {};
  const treatmentLogs = payload.treatmentLogs && Array.isArray(payload.treatmentLogs.logs) ? payload.treatmentLogs.logs : [];
  const now = dashboardCoerceDate_(payload.now) || new Date();
  const tz = dashboardResolveTimeZone_();
  const userEmail = dashboardNormalizeEmail_(payload.user || '');
  const isAdmin = typeof isAdminUser_ === 'function' ? !!isAdminUser_() : false;

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

  const targetScope = resolveDashboardOverviewTargets_({
    userEmail,
    isAdmin,
    patientInfo,
    patients,
    treatmentLogs,
    now,
    tz
  });

  const invoiceUnconfirmed = buildOverviewFromTasks_(tasks, ['invoiceUnconfirmed'], targetScope, patientNameMap);
  const consentRelated = buildOverviewFromConsent_(tasks, patientInfo, targetScope, patientNameMap);
  const visitSummary = buildOverviewFromVisits_(visits, targetScope, patientNameMap, now, tz);

  return {
    invoiceUnconfirmed,
    consentRelated,
    visitSummary
  };
}

function resolveDashboardOverviewTargets_(params) {
  const payload = params || {};
  const isAdmin = !!payload.isAdmin;
  const userEmail = dashboardNormalizeEmail_(payload.userEmail || '');
  const patients = Array.isArray(payload.patients) ? payload.patients : [];
  const patientInfo = payload.patientInfo || {};
  const treatmentLogs = Array.isArray(payload.treatmentLogs) ? payload.treatmentLogs : [];
  const now = dashboardCoerceDate_(payload.now) || new Date();
  const tz = payload.tz || dashboardResolveTimeZone_();
  const monthStart = dashboardStartOfMonth_(tz, now);
  const allowed = new Set();

  const addAllPatients = () => {
    patients.forEach(entry => {
      if (entry && entry.patientId) allowed.add(entry.patientId);
    });
    Object.keys(patientInfo || {}).forEach(pid => {
      if (pid) allowed.add(pid);
    });
  };

  if (isAdmin || !userEmail) {
    addAllPatients();
    return { patientIds: allowed, applyFilter: false };
  }

  treatmentLogs.forEach(entry => {
    if (!entry || !entry.patientId || !entry.timestamp) return;
    const ts = dashboardCoerceDate_(entry.timestamp);
    if (!ts || ts < monthStart) return;
    const email = dashboardNormalizeEmail_(entry.createdByEmail || '');
    if (!email || email !== userEmail) return;
    allowed.add(entry.patientId);
  });

  return { patientIds: allowed, applyFilter: true };
}

function buildOverviewFromTasks_(tasks, targetTypes, scope, patientNameMap) {
  const typeSet = new Set(targetTypes || []);
  const items = [];
  const seen = new Set();
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;

  (tasks || []).forEach(task => {
    const pid = task && task.patientId ? String(task.patientId).trim() : '';
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    const type = task && task.type ? String(task.type) : '';
    if (!typeSet.has(type)) return;
    if (seen.has(pid)) return;
    seen.add(pid);
    items.push({ patientId: pid, name: task.name || patientNameMap[pid] || '' });
  });

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  return { count: items.length, items };
}

function buildOverviewFromConsent_(tasks, patientInfo, scope, patientNameMap) {
  const items = [];
  const seen = new Set();
  const taskItems = buildOverviewFromTasks_(tasks, ['consentExpired', 'consentWarning'], scope, patientNameMap);
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;
  (taskItems.items || []).forEach(entry => {
    if (!entry.patientId || seen.has(entry.patientId)) return;
    seen.add(entry.patientId);
    items.push(entry);
  });

  Object.keys(patientInfo || {}).forEach(pid => {
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    if (seen.has(pid)) return;
    const info = patientInfo[pid] || {};
    const consentValue = info.consentExpiry || (info.raw && (info.raw['同意期限'] || info.raw['同意有効期限'])) || '';
    if (String(consentValue || '').trim()) return;
    seen.add(pid);
    items.push({ patientId: pid, name: info.name || patientNameMap[pid] || '' });
  });

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  return { count: items.length, items };
}

function buildOverviewFromVisits_(visits, scope, patientNameMap, now, tz) {
  const targetNow = dashboardCoerceDate_(now) || new Date();
  const todayKey = dashboardFormatDate_(targetNow, tz, 'yyyy-MM-dd');
  const yesterday = new Date(targetNow.getTime() - 24 * 60 * 60 * 1000);
  const yesterdayKey = dashboardFormatDate_(yesterday, tz, 'yyyy-MM-dd');
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;
  const byDate = {
    [todayKey]: { count: 0, items: [] },
    [yesterdayKey]: { count: 0, items: [] }
  };
  const seenByDate = {
    [todayKey]: new Set(),
    [yesterdayKey]: new Set()
  };

  (visits || []).forEach(entry => {
    const pid = entry && entry.patientId ? String(entry.patientId).trim() : '';
    if (!pid || (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid))) return;
    const dateKey = entry.dateKey || '';
    if (!Object.prototype.hasOwnProperty.call(byDate, dateKey)) return;
    if (seenByDate[dateKey].has(pid)) return;
    seenByDate[dateKey].add(pid);
    byDate[dateKey].count += 1;
    const name = entry.patientName || patientNameMap[pid] || '';
    byDate[dateKey].items.push({ patientId: pid, name });
  });

  Object.keys(byDate).forEach(key => {
    byDate[key].items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  });

  return {
    today: byDate[todayKey] || { count: 0, items: [] },
    yesterday: byDate[yesterdayKey] || { count: 0, items: [] }
  };
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

  (treatmentLogs || []).forEach(entry => {
    if (!entry || !entry.timestamp) return;
    const ts = dashboardCoerceDate_(entry.timestamp);
    if (!ts) return;
    const dateKey = entry.dateKey || dashboardFormatDate_(ts, tz, 'yyyy-MM-dd');
    const inMonthlyScope = ts >= monthStart;
    const inDailyScope = dateKey === todayKey || dateKey === yesterdayKey;
    if (!inMonthlyScope && !inDailyScope) return;
    const pid = dashboardNormalizePatientId_(entry.patientId);
    if (!pid || !Object.prototype.hasOwnProperty.call(patientInfo, pid)) return;
    allowed.add(pid);
  });

  return { patientIds: allowed };
}

function filterDashboardItemsByPatientId_(items, allowedPatientIds) {
  if (!Array.isArray(items)) return [];
  if (!allowedPatientIds || !(allowedPatientIds instanceof Set)) return items.slice();
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
