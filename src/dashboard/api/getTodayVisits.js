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
  const setupIncomplete = !!(treatment && treatment.setupIncomplete)
    || !!(notesResult && notesResult.setupIncomplete);

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

  return { visits, warnings, setupIncomplete };
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
