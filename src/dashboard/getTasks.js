/**
 * ダッシュボード上で表示するタスクを抽出する。
 * @param {Object} [options]
 * @param {Object} [options.patientInfo] - loadPatientInfo() の戻り値を差し替える際に利用。
 * @param {Object} [options.notes] - loadNotes() の戻り値を差し替える際に利用。
 * @param {Object} [options.aiReports] - loadAIReports() の戻り値を差し替える際に利用。
 * @param {Object<string, boolean>} [options.invoiceConfirmations] - 請求書確認フラグを患者ID単位で差し込む。
 * @param {Date} [options.now] - テスト用に現在日時を差し替え。
 * @return {{tasks: Object[], warnings: string[]}}
 */
function getTasks(options) {
  const opts = options || {};
  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();

  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const notesResult = opts.notes || (typeof loadNotes === 'function' ? loadNotes() : null);
  const aiReports = opts.aiReports || (typeof loadAIReports === 'function' ? loadAIReports() : null);
  const invoiceConfirmations = opts.invoiceConfirmations || {};

  const warnings = [];
  if (patientInfo && Array.isArray(patientInfo.warnings)) warnings.push.apply(warnings, patientInfo.warnings);
  if (notesResult && Array.isArray(notesResult.warnings)) warnings.push.apply(warnings, notesResult.warnings);
  if (aiReports && Array.isArray(aiReports.warnings)) warnings.push.apply(warnings, aiReports.warnings);

  const notesByPatient = notesResult && notesResult.notes ? notesResult.notes : {};
  const reportsByPatient = aiReports && aiReports.reports ? aiReports.reports : {};

  const tasks = [];
  Object.keys(patients).forEach(pid => {
    const patient = patients[pid] || {};
    const normalized = dashboardNormalizePatientId_(pid);
    if (!normalized) return;
    const name = patient.name || patient.patientName || '';

    // 同意書期限
    const consentDate = dashboardParseTimestamp_(patient.consentExpiry || (patient.raw && patient.raw['同意期限']));
    if (consentDate) {
      const daysUntil = daysBetween_(now, consentDate, true);
      if (daysUntil <= -180) {
        tasks.push(makeTask_(normalized, name, 'consentExpired', 'critical', dashboardFormatDate_(consentDate, tz, 'yyyy-MM-dd')));
      } else if (daysUntil <= 14) {
        tasks.push(makeTask_(normalized, name, 'consentWarning', 'warning', dashboardFormatDate_(consentDate, tz, 'yyyy-MM-dd')));
      }
    }

    // 申し送り遅延
    const note = notesByPatient[normalized];
    const noteTs = note ? (note.timestamp instanceof Date ? note.timestamp : dashboardParseTimestamp_(note.timestamp || note.when)) : null;
    if (!noteTs || daysBetween_(noteTs, now) >= 30) {
      tasks.push(makeTask_(normalized, name, 'handoverDelayed', 'warning', noteTs ? dashboardFormatDate_(noteTs, tz, 'yyyy-MM-dd') : '未入力'));
    }

    // 医師報告書遅延
    const reportRaw = reportsByPatient[normalized];
    const reportDate = reportRaw ? dashboardParseTimestamp_(reportRaw) : null;
    if (!reportDate || daysBetween_(reportDate, now) >= 180) {
      tasks.push(makeTask_(normalized, name, 'aiReportDelayed', 'warning', reportRaw || '未発行'));
    }

    // 請求書確認未完了
    if (Object.prototype.hasOwnProperty.call(invoiceConfirmations, normalized) && !invoiceConfirmations[normalized]) {
      tasks.push(makeTask_(normalized, name, 'invoiceUnconfirmed', 'warning', '請求書確認未完了'));
    }
  });

  return { tasks, warnings };
}

function makeTask_(patientId, name, type, severity, detail) {
  return { patientId, name, type, severity, detail };
}

function daysBetween_(from, to, futurePositive) {
  const start = dashboardCoerceDate_(from);
  const end = dashboardCoerceDate_(to);
  if (!start || !end) return Number.POSITIVE_INFINITY;
  const diff = end.getTime() - start.getTime();
  const days = Math.floor(diff / (24 * 60 * 60 * 1000));
  return futurePositive ? days : Math.abs(days);
}

if (typeof dashboardResolveTimeZone_ === 'undefined') {
  function dashboardResolveTimeZone_() {
    if (typeof Session !== 'undefined' && Session && typeof Session.getScriptTimeZone === 'function') {
      const tz = Session.getScriptTimeZone();
      if (tz) return tz;
    }
    return 'Asia/Tokyo';
  }
}

if (typeof dashboardFormatDate_ === 'undefined') {
  function dashboardFormatDate_(date, tz, format) {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.formatDate === 'function') {
      try { return Utilities.formatDate(date, tz, format); } catch (e) { /* ignore */ }
    }
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
    const iso = date.toISOString();
    if (format && format.indexOf('HH') >= 0) return iso.replace('T', ' ').slice(0, 16);
    if (format && format.indexOf('-') >= 0) return iso.slice(0, 10);
    return iso;
  }
}

if (typeof dashboardCoerceDate_ === 'undefined') {
  function dashboardCoerceDate_(value) {
    if (value instanceof Date) return value;
    if (value && typeof value.getTime === 'function') {
      const ts = value.getTime();
      if (Number.isFinite(ts)) return new Date(ts);
    }
    if (value === null || value === undefined) return null;
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
}

if (typeof dashboardParseTimestamp_ === 'undefined') {
  function dashboardParseTimestamp_(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number' && Number.isFinite(value)) return new Date(value);
    const str = String(value == null ? '' : value).trim();
    if (!str) return null;
    const parsed = new Date(str);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
}

if (typeof dashboardNormalizePatientId_ === 'undefined') {
  function dashboardNormalizePatientId_(value) {
    const raw = value == null ? '' : value;
    return String(raw).trim();
  }
}

if (typeof loadPatientInfo === 'undefined') {
  function loadPatientInfo() { return { patients: {}, nameToId: {}, warnings: [] }; }
}

if (typeof loadNotes === 'undefined') {
  function loadNotes() { return { notes: {}, warnings: [] }; }
}

if (typeof loadAIReports === 'undefined') {
  function loadAIReports() { return { reports: {}, warnings: [] }; }
}
