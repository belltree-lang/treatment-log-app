/**
 * 今日の施術実績と、今日より前で最も新しい1日分の施術実績をタイムライン形式で返す。
 * @param {Object} [options]
 * @param {Object} [options.treatmentLogs] - loadTreatmentLogs() の戻り値を差し替える際に利用。
 * @param {Object} [options.patientInfo] - loadPatientInfo() の戻り値を差し替える際に利用。
 * @param {Object} [options.notes] - loadNotes() の戻り値を差し替える際に利用。
 * @param {Date} [options.now] - テスト用に現在日時を差し替え。
 * @param {Set<string>} [options.visiblePatientIds] - 表示対象患者ID。null の場合は全件。
 * @return {{visits: Object[], warnings: string[]}}
 */
function getTodayVisits(options) {
  const opts = options || {};
  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const visiblePatientIds = opts.visiblePatientIds && typeof opts.visiblePatientIds.has === 'function' ? opts.visiblePatientIds : null;
  const todayKey = dashboardFormatDate_(now, tz, 'yyyy-MM-dd');

  const treatment = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function' ? loadTreatmentLogs() : null);
  const logs = treatment && Array.isArray(treatment.logs) ? treatment.logs : [];
  const patientInfo = opts.patientInfo && opts.patientInfo.patients ? opts.patientInfo.patients : {};
  const notesResult = opts.notes || (typeof loadNotes === 'function' ? loadNotes() : null);
  const notes = notesResult && notesResult.notes ? notesResult.notes : {};

  const warnings = [];
  if (treatment && Array.isArray(treatment.warnings)) warnings.push.apply(warnings, treatment.warnings);
  if (notesResult && Array.isArray(notesResult.warnings)) warnings.push.apply(warnings, notesResult.warnings);
  const setupIncomplete = !!(treatment && treatment.setupIncomplete)
    || !!(notesResult && notesResult.setupIncomplete);
  if (typeof dashboardLogContext_ === 'function') {
    dashboardLogContext_('getTodayVisits:setupIncomplete', `treatmentLogs=${!!(treatment && treatment.setupIncomplete)} notes=${!!(notesResult && notesResult.setupIncomplete)}`);
  } else if (typeof dashboardWarn_ === 'function') {
    dashboardWarn_(`[getTodayVisits:setupIncomplete] treatmentLogs=${!!(treatment && treatment.setupIncomplete)} notes=${!!(notesResult && notesResult.setupIncomplete)}`);
  }

  const normalizedVisits = [];
  logs.forEach(entry => {
    if (!entry || !entry.timestamp) return;
    const ts = dashboardCoerceDate_(entry.timestamp);
    if (!ts) return;
    const dateKey = entry.dateKey || dashboardFormatDate_(ts, tz, 'yyyy-MM-dd');

    const patientId = dashboardNormalizePatientId_(entry.patientId);
    const master = patientId && Object.prototype.hasOwnProperty.call(patientInfo, patientId)
      ? patientInfo[patientId]
      : null;
    const patientName = entry.patientName || (master && (master.name || master.patientName)) || '';
    const time = dashboardFormatDate_(ts, tz, 'HH:mm');
    const noteStatus = resolveHandoverStatus_(dateKey, patientId, notes, tz);

    if (!patientId) return;
    if (visiblePatientIds && !visiblePatientIds.has(patientId)) return;

    normalizedVisits.push({ patientId, patientName, time, dateKey, noteStatus });
  });

  const latestPastDate = normalizedVisits
    .filter(visit => visit.dateKey < todayKey)
    .map(visit => visit.dateKey)
    .sort((a, b) => b.localeCompare(a))[0] || '';

  const visits = normalizedVisits.filter(visit => visit.dateKey === todayKey || (latestPastDate && visit.dateKey === latestPastDate));

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
