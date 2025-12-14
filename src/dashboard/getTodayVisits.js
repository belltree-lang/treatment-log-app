/**
 * 今日と昨日の施術実績をタイムライン形式で返す。
 * @param {Object} [options]
 * @param {Object} [options.treatmentLogs] - loadTreatmentLogs() の戻り値を差し替える際に利用。
 * @param {Object} [options.notes] - loadNotes() の戻り値を差し替える際に利用。
 * @param {Date} [options.now] - テスト用に現在日時を差し替え。
 * @return {{visits: Object[], warnings: string[]}}
 */
function getTodayVisits(options) {
  const opts = options || {};
  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const todayKey = dashboardFormatDate_(now, tz, 'yyyy-MM-dd');
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const yesterdayKey = dashboardFormatDate_(yesterday, tz, 'yyyy-MM-dd');

  const treatment = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function' ? loadTreatmentLogs() : null);
  const logs = treatment && Array.isArray(treatment.logs) ? treatment.logs : [];
  const notesResult = opts.notes || (typeof loadNotes === 'function' ? loadNotes() : null);
  const notes = notesResult && notesResult.notes ? notesResult.notes : {};

  const warnings = [];
  if (treatment && Array.isArray(treatment.warnings)) warnings.push.apply(warnings, treatment.warnings);
  if (notesResult && Array.isArray(notesResult.warnings)) warnings.push.apply(warnings, notesResult.warnings);

  const visits = [];
  logs.forEach(entry => {
    if (!entry || !entry.timestamp) return;
    const ts = dashboardCoerceDate_(entry.timestamp);
    if (!ts) return;
    const dateKey = entry.dateKey || dashboardFormatDate_(ts, tz, 'yyyy-MM-dd');
    if (dateKey !== todayKey && dateKey !== yesterdayKey) return;

    const patientId = dashboardNormalizePatientId_(entry.patientId);
    const patientName = entry.patientName || '';
    const time = dashboardFormatDate_(ts, tz, 'HH:mm');
    const noteStatus = resolveHandoverStatus_(dateKey, patientId, notes, tz);

    visits.push({ patientId, patientName, time, dateKey, noteStatus });
  });

  visits.sort((a, b) => {
    if (a.dateKey === b.dateKey) return a.time.localeCompare(b.time);
    return a.dateKey.localeCompare(b.dateKey);
  });

  return { visits, warnings };
}

function resolveHandoverStatus_(dateKey, patientId, notes, tz) {
  const note = notes && patientId ? notes[patientId] : null;
  if (note) {
    const ts = note.timestamp instanceof Date ? note.timestamp : dashboardParseTimestamp_(note.timestamp || note.when);
    const noteDate = ts ? dashboardFormatDate_(ts, tz, 'yyyy-MM-dd') : '';
    if (noteDate && noteDate === dateKey) return '◎';
  }
  return '△';
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
    if (format === 'HH:mm') return iso.slice(11, 16);
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

if (typeof loadTreatmentLogs === 'undefined') {
  function loadTreatmentLogs() { return { logs: [], warnings: [] }; }
}

if (typeof loadNotes === 'undefined') {
  function loadNotes() { return { notes: {}, warnings: [] }; }
}
