/**
 * ダッシュボードの主要データをまとめて取得し、JSON 形式で返す。
 * エラーが発生した場合は meta.error にメッセージを格納する。
 * @param {Object} [options]
 * @return {{tasks: Object[], todayVisits: Object[], patients: Object[], warnings: string[], meta: Object}}
 */
function getDashboardData(options) {
  const opts = options || {};
  if (opts.mock && typeof buildDashboardMockData_ === 'function') {
    const mockOptions = buildDashboardMockData_(opts) || {};
    const normalized = Object.assign({}, mockOptions);
    normalized.mock = '';
    return getDashboardData(normalized);
  }
  const cacheOptions = opts.cache === false ? { cache: false } : {};
  const meta = {
    generatedAt: dashboardFormatDate_(new Date(), dashboardResolveTimeZone_(), "yyyy-MM-dd'T'HH:mm:ssXXX") || new Date().toISOString(),
    user: opts.user || dashboardResolveUser_(),
    setupIncomplete: false
  };

  try {
    const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo(cacheOptions) : { patients: {}, nameToId: {}, warnings: [] });
    const notes = opts.notes || (typeof loadNotes === 'function' ? loadNotes(Object.assign({ email: meta.user }, cacheOptions)) : { notes: {}, warnings: [] });
    const aiReports = opts.aiReports || (typeof loadAIReports === 'function' ? loadAIReports(cacheOptions) : { reports: {}, warnings: [] });
    const invoices = opts.invoices || (typeof loadInvoices === 'function'
      ? loadInvoices(Object.assign({ patientInfo, now: opts.now }, cacheOptions))
      : { invoices: {}, warnings: [] });
    const treatmentLogs = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function'
      ? loadTreatmentLogs(Object.assign({ patientInfo, now: opts.now }, cacheOptions))
      : { logs: [], warnings: [] });
    const responsible = opts.responsible || (typeof assignResponsibleStaff === 'function'
      ? assignResponsibleStaff(Object.assign({ patientInfo, treatmentLogs, now: opts.now }, cacheOptions))
      : { responsible: {}, warnings: [] });

    const tasksResult = opts.tasksResult || (typeof getTasks === 'function' ? getTasks({
      patientInfo,
      notes,
      aiReports,
      invoiceConfirmations: opts.invoiceConfirmations,
      now: opts.now
    }) : { tasks: [], warnings: [] });

    const visitsResult = opts.visitsResult || (typeof getTodayVisits === 'function' ? getTodayVisits({
      treatmentLogs,
      notes,
      now: opts.now
    }) : { visits: [], warnings: [] });

    const patients = buildDashboardPatients_(patientInfo, {
      notes,
      aiReports,
      invoices,
      responsible,
      treatmentLogs
    });

    const warningState = collectDashboardWarnings_([
      patientInfo,
      notes,
      aiReports,
      invoices,
      treatmentLogs,
      responsible,
      tasksResult,
      visitsResult
    ]);

    meta.setupIncomplete = warningState.setupIncomplete;

    return {
      tasks: tasksResult && tasksResult.tasks ? tasksResult.tasks : [],
      todayVisits: visitsResult && visitsResult.visits ? visitsResult.visits : [],
      patients,
      warnings: warningState.warnings,
      meta
    };
  } catch (err) {
    meta.error = err && err.message ? err.message : String(err);
    return { tasks: [], todayVisits: [], patients: [], warnings: [], meta };
  }
}

function buildDashboardPatients_(patientInfo, sources) {
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

  const additionalIds = new Set([
    ...Object.keys(notes || {}),
    ...Object.keys(aiReports || {}),
    ...Object.keys(invoices || {}),
    ...Object.keys(responsible || {})
  ]);
  additionalIds.forEach(pid => {
    if (!pid || Object.prototype.hasOwnProperty.call(basePatients, pid)) return;
    addPatient(pid, {});
  });

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

function collectDashboardWarnings_(results) {
  const warnings = [];
  const seen = new Set();
  let setupIncomplete = false;

  (results || []).forEach(entry => {
    if (entry && entry.setupIncomplete) setupIncomplete = true;
    if (entry && Array.isArray(entry.warnings)) {
      entry.warnings.forEach(warning => {
        const normalized = normalizeDashboardWarning_(warning);
        if (!normalized || seen.has(normalized)) return;
        seen.add(normalized);
        warnings.push(normalized);
        if (isDashboardSetupIncompleteWarning_(normalized)) setupIncomplete = true;
      });
    }
  });

  return { warnings, setupIncomplete };
}

function normalizeDashboardWarning_(warning) {
  return String(warning == null ? '' : warning).trim();
}

function isDashboardSetupIncompleteWarning_(warning) {
  const normalized = normalizeDashboardWarning_(warning);
  if (!normalized) return false;
  if (normalized.indexOf('シートが見つかりません') >= 0) return true;
  if (normalized.indexOf('スプレッドシートを取得できませんでした') >= 0) return true;
  if (normalized.indexOf('請求書フォルダが取得できませんでした') >= 0) return true;
  return false;
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
