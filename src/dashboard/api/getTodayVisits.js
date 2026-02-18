/**
 * 今日の施術実績と、今日より前で最も新しい1日分の施術実績をタイムライン形式で返す。
 * @param {Object} [options]
 * @param {Object} [options.treatmentLogs] - loadTreatmentLogs() の戻り値を差し替える際に利用。
 * @param {Object} [options.patientInfo] - loadPatientInfo() の戻り値を差し替える際に利用。
 * @param {Object} [options.notes] - loadNotes() の戻り値を差し替える際に利用。
 * @param {Date} [options.now] - テスト用に現在日時を差し替え。
 * @return {{today: {date: string, visits: Object[]}, previous: {date: (string|null), visits: Object[]}, warnings: string[]}}
 */
function getTodayVisits(options) {
  const opts = options || {};
  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const todayKey = dashboardFormatDate_(now, tz, 'yyyy-MM-dd');

  const treatment = opts.treatmentLogs || (typeof loadTreatmentLogs === 'function' ? loadTreatmentLogs() : null);
  const logs = treatment && Array.isArray(treatment.logs) ? treatment.logs : [];
  const patientInfo = opts.patientInfo && opts.patientInfo.patients ? opts.patientInfo.patients : {};
  const notesResult = opts.notes || (typeof loadNotes === 'function' ? loadNotes() : null);
  const notes = notesResult && notesResult.notes ? notesResult.notes : {};

  if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
    Logger.log('[TODAY VISITS ENTRY] logs length=' + (logs ? logs.length : 'null'));
  }

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

  const visitsByDate = {};
  let previousKey = '';
  logs.forEach(entry => {
    if (!entry || !entry.timestamp) return;
    const ts = dashboardCoerceDate_(entry.timestamp);
    if (!ts) return;
    const dateKey = String(entry.dateKey || '').trim() || dashboardFormatDate_(ts, tz, 'yyyy-MM-dd');

    const patientId = dashboardNormalizePatientId_(entry.patientId);
    const master = patientId && Object.prototype.hasOwnProperty.call(patientInfo, patientId)
      ? patientInfo[patientId]
      : null;
    const patientName = entry.patientName || (master && (master.name || master.patientName)) || '';
    const time = dashboardFormatDate_(ts, tz, 'HH:mm');
    const noteStatus = resolveHandoverStatus_(dateKey, patientId, notes, tz);

    if (!patientId) return;
    if (!Object.prototype.hasOwnProperty.call(visitsByDate, dateKey)) visitsByDate[dateKey] = [];
    visitsByDate[dateKey].push({ patientId, patientName, time, dateKey, noteStatus });

    if (dateKey < todayKey && (!previousKey || dateKey > previousKey)) previousKey = dateKey;
  });

  const todayVisits = Object.prototype.hasOwnProperty.call(visitsByDate, todayKey)
    ? visitsByDate[todayKey]
    : [];
  const previousVisits = previousKey && Object.prototype.hasOwnProperty.call(visitsByDate, previousKey)
    ? visitsByDate[previousKey]
    : [];
  const sortByTime = (a, b) => a.time.localeCompare(b.time);
  todayVisits.sort(sortByTime);
  previousVisits.sort(sortByTime);

  const result = {
    today: {
      date: todayKey,
      visits: todayVisits
    },
    previous: {
      date: previousKey || null,
      visits: previousVisits
    }
  };
  if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
    Logger.log('[TODAY VISITS RETURN] today=' + result.today.visits.length + ' previous=' + result.previous.visits.length);
  }

  return {
    today: result.today,
    previous: result.previous,
    warnings,
    setupIncomplete
  };
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
