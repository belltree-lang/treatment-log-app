/**
 * ダッシュボードの主要データをまとめて取得し、JSON 形式で返す。
 * エラーが発生した場合は meta.error にメッセージを格納する。
 *
 * @param {Object} [options]
 * @param {Object} [options.patientInfo] - 事前に取得した患者情報を差し込む場合に利用。
 * @param {Object} [options.notes] - 申し送りデータを差し込む場合に利用。
 * @param {Object} [options.aiReports] - AI報告書データを差し込む場合に利用。
 * @param {Object} [options.invoices] - 請求書リンクデータを差し込む場合に利用。
 * @param {Object} [options.treatmentLogs] - 施術録データを差し込む場合に利用。
 * @param {Object} [options.responsible] - 担当者マッピングを差し込む場合に利用。
 * @param {Object} [options.tasksResult] - タスク抽出結果を差し込む場合に利用。
 * @param {Object} [options.visitsResult] - 今日の訪問抽出結果を差し込む場合に利用。
 * @param {Object<string, boolean>} [options.invoiceConfirmations] - 請求書確認フラグを直接指定する場合に利用。
 * @param {Date} [options.now] - テスト用に現在日時を差し替え。
 * @param {string} [options.user] - セッションユーザーを明示的に指定する場合に利用。
 * @return {{tasks: Object[], todayVisits: Object[], patients: Object[], warnings: string[], meta: Object}}
 */
function getDashboardData(options) {
  const opts = options || {};
  const meta = {
    generatedAt: dashboardFormatDate_(new Date(), dashboardResolveTimeZone_(), "yyyy-MM-dd'T'HH:mm:ssXXX") || new Date().toISOString(),
    user: opts.user || dashboardResolveUser_()
  };

  try {
    const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : { patients: {}, nameToId: {}, warnings: [] });
    const notes = opts.notes || (typeof loadNotes === 'function' ? loadNotes({ email: meta.user }) : { notes: {}, warnings: [] });
    const aiReports = opts.aiReports || (typeof loadAIReports === 'function' ? loadAIReports() : { reports: {}, warnings: [] });
    const invoices = opts.invoices || (typeof loadInvoices === 'function' ? loadInvoices({ patientInfo }) : { invoices: {}, warnings: [] });
    const treatmentLogs = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function' ? loadTreatmentLogs({ patientInfo }) : { logs: [], warnings: [] });
    const responsible = opts.responsible || (typeof assignResponsibleStaff === 'function' ? assignResponsibleStaff({ patientInfo, treatmentLogs }) : { responsible: {}, warnings: [] });

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

    const warnings = collectDashboardWarnings_([
      patientInfo,
      notes,
      aiReports,
      invoices,
      treatmentLogs,
      responsible,
      tasksResult,
      visitsResult
    ]);

    return {
      tasks: tasksResult && tasksResult.tasks ? tasksResult.tasks : [],
      todayVisits: visitsResult && visitsResult.visits ? visitsResult.visits : [],
      patients,
      warnings,
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
  (results || []).forEach(entry => {
    if (entry && Array.isArray(entry.warnings)) {
      warnings.push.apply(warnings, entry.warnings);
    }
  });
  return warnings;
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
