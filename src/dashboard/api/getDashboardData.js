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
  const logCachePerf = (status, key, details) => {
    const suffix = details ? ` ${details}` : '';
    logPerf(`[perf] ${status} key=${key}${suffix}`);
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
    const treatmentLogs = measureStep('loadTreatmentLogs', () => {
      if (opts.treatmentLogs) return opts.treatmentLogs;
      if (typeof loadTreatmentLogs !== 'function') return { logs: [], warnings: [] };

      const now = dashboardCoerceDate_(opts.now) || new Date();
      const rawMonthKey = dashboardFormatDate_(now, dashboardResolveTimeZone_(), 'yyyyMM');
      const normalizedMonthKey = String(rawMonthKey || '').replace(/[^0-9]/g, '').slice(0, 6);
      const monthKey = normalizedMonthKey || `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}`;
      const cacheKey = `dashboard:treatmentLogs:${monthKey}`;
      return dashboardGetTreatmentLogsFromCache_(cacheKey, monthKey, () => loadTreatmentLogs({
        patientInfo,
        now: opts.now,
        cache: opts.cache,
        dashboardSpreadsheet
      }), logCachePerf);
    });
    const totalTreatmentLogs = treatmentLogs && treatmentLogs.logs ? treatmentLogs.logs.length : 0;
    logContext('getDashboardData:loadTreatmentLogs', `logs=${totalTreatmentLogs} warnings=${(treatmentLogs && treatmentLogs.warnings ? treatmentLogs.warnings.length : 0)} setupIncomplete=${!!(treatmentLogs && treatmentLogs.setupIncomplete)}`);
    if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
      Logger.log(JSON.stringify({
        patientCount: Object.keys(patientInfo && patientInfo.patients ? patientInfo.patients : {}).length,
        noteCount: Object.keys(notes && notes.notes ? notes.notes : {}).length,
        reportCount: Object.keys(aiReports && aiReports.reports ? aiReports.reports : {}).length,
        invoiceCount: Object.keys(invoices && invoices.invoices ? invoices.invoices : {}).length,
        treatmentLogCount: totalTreatmentLogs
      }));
    }
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

    if (!isAdmin) {
      const patientMasterIds = Object.keys(patientMaster || {});
      let consentEligiblePatients = 0;
      let consentEligibleButOutOfScope = 0;
      let parseFailedCount = 0;
      let consentAcquiredExcludedCount = 0;
      let scopeExcludedCount = 0;

      patientMasterIds.forEach(pid => {
        const info = patientMaster[pid] || {};
        const consentExpiryResolved = resolveConsentExpiry_(info);
        const consentExpiryDate = parseConsentDate_(consentExpiryResolved.value);
        const consentAcquired = dashboardIsConsentAcquired_(info.raw);
        const inVisibleScope = visiblePatientIds.has(pid);
        const hasConsentExpiry = consentExpiryResolved.value != null && String(consentExpiryResolved.value).trim() !== '';

        if (hasConsentExpiry && !consentExpiryDate) parseFailedCount += 1;
        if (consentAcquired) consentAcquiredExcludedCount += 1;
        if (!inVisibleScope) scopeExcludedCount += 1;

        logContext('getDashboardData:consentEligibilityPatient', JSON.stringify({
          pid,
          hasConsentExpiry,
          parsedConsentExpiry: consentExpiryDate ? consentExpiryDate.toISOString() : null,
          consentAcquired,
          inVisibleScope
        }));

        const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        const parsedConsentExpiry = consentExpiryDate
          ? new Date(consentExpiryDate.getFullYear(), consentExpiryDate.getMonth(), consentExpiryDate.getDate())
          : null;
        const diffMs = parsedConsentExpiry ? parsedConsentExpiry.getTime() - today.getTime() : null;
        const diffDays = diffMs == null ? null : Math.floor(diffMs / (1000 * 60 * 60 * 24));
        const threshold = null;
        const finalCondition = Boolean(consentExpiryDate) && !consentAcquired;
        logContext('consent-eligible-debug', JSON.stringify({
          pid,
          hasConsentExpiry,
          parsedConsentExpiry: parsedConsentExpiry ? parsedConsentExpiry.toISOString() : null,
          consentAcquired,
          inVisibleScope,
          todayISO: today.toISOString(),
          expiryISO: parsedConsentExpiry ? parsedConsentExpiry.toISOString() : null,
          diffMs,
          diffDays,
          threshold,
          finalCondition
        }));

        if (!consentExpiryDate) return;
        if (consentAcquired) return;
        consentEligiblePatients += 1;
        if (!inVisibleScope) consentEligibleButOutOfScope += 1;
      });

      logContext('getDashboardData:consentEligibleFormula', 'consentEligiblePatients = count(pid where parseConsentDate_(resolveConsentExpiry_(patient).value) != null && dashboardIsConsentAcquired_(patient.raw) === false)');
      logContext('getDashboardData:consentScopeMetrics', JSON.stringify({
        totalPatients: patientMasterIds.length,
        consentEligiblePatients,
        visiblePatientIdsSize: visiblePatientIds.size,
        consentEligibleButOutOfScope,
        parseFailedCount,
        consentAcquiredExcludedCount,
        scopeExcludedCount
      }));
      logContext('getDashboardData:visibleVsEligible', JSON.stringify({
        visiblePatientIdsSize: visiblePatientIds.size,
        consentEligiblePatients,
        diff: visiblePatientIds.size - consentEligiblePatients
      }));
      logContext(
        'getDashboardData:consentMissingByRecentLog',
        `staffOnly consent期限あり&同意取得確認なし&直近50日ログなし=${consentEligibleButOutOfScope}`
      );
      logContext(
        'getDashboardData:visibleScopeRoutes',
        'overview.consentRelated=buildDashboardOverview_ -> buildOverviewFromConsent_(allowedPatientIds=visiblePatientIds), patients/statusTags=buildDashboardPatients_(allowedPatientIds=visiblePatientIds) -> buildDashboardPatientStatusTags_'
      );
    }

    const unpaidAlertsResult = measureStep('loadUnpaidAlerts', () => (opts.unpaidAlerts || (typeof loadUnpaidAlerts === 'function'
      ? loadUnpaidAlerts({ patientInfo, now: opts.now, cache: opts.cache, dashboardSpreadsheet, visiblePatientIds })
      : { alerts: [], warnings: [] })));
    logContext('getDashboardData:loadUnpaidAlerts', `alerts=${(unpaidAlertsResult && unpaidAlertsResult.alerts ? unpaidAlertsResult.alerts.length : 0)} warnings=${(unpaidAlertsResult && unpaidAlertsResult.warnings ? unpaidAlertsResult.warnings.length : 0)} setupIncomplete=${!!(unpaidAlertsResult && unpaidAlertsResult.setupIncomplete)}`);

    const tasksResult = { tasks: [], warnings: [] };

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
      treatmentLogs,
      now: opts.now
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
      visitsResult
    ]);

    meta.setupIncomplete = warningState.setupIncomplete;
    logContext('getDashboardData:setupIncomplete', `result=${meta.setupIncomplete} warnings=${warningState.warnings.length}`);

    const overview = buildDashboardOverview_({
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
    const invoiceUnconfirmed = overview && overview.invoiceUnconfirmed ? overview.invoiceUnconfirmed : { items: [] };
    logContext('getDashboardData:overviewInvoiceUnconfirmed', `items.length=${Array.isArray(invoiceUnconfirmed.items) ? invoiceUnconfirmed.items.length : 0}`);
    return {
      tasks: [],
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

function dashboardGetTreatmentLogsFromCache_(cacheKey, monthKey, loader, logCachePerf) {
  const shouldLog = typeof logCachePerf === 'function';
  const emit = (status, details) => {
    if (shouldLog) logCachePerf(status, cacheKey, details);
  };
  const canUseCacheService = typeof CacheService !== 'undefined'
    && CacheService
    && typeof CacheService.getScriptCache === 'function';
  const cacheTtlSeconds = 300;

  if (canUseCacheService) {
    let cache = null;
    try {
      cache = CacheService.getScriptCache();
      const cachedRaw = cache.get(cacheKey);
      if (cachedRaw) {
        const parsed = JSON.parse(cachedRaw);
        if (parsed && parsed.monthKey === monthKey && parsed.payload) {
          emit('cacheHit', 'source=CacheService');
          return reviveTreatmentLogsCachePayload_(parsed.payload);
        }
      }
      emit('cacheMiss', 'source=CacheService');
    } catch (err) {
      emit('cacheMiss', `reason=cacheError error=${err && err.message ? err.message : err}`);
    }

    const loaded = loader();

    if (cache) {
      try {
        const payload = JSON.stringify({
          monthKey,
          payload: sanitizeTreatmentLogsCachePayload_(loaded)
        });
        if (payload.length > 90000) {
          emit('cacheSkip', `reason=tooLarge size=${payload.length}`);
          return loaded;
        }
        cache.put(cacheKey, payload, cacheTtlSeconds);
      } catch (err) {
        emit('cacheSkip', `reason=cacheError error=${err && err.message ? err.message : err}`);
      }
    }

    return loaded;
  }

  emit('cacheMiss', 'reason=cacheUnavailable');
  return loader();
}

function sanitizeTreatmentLogsCachePayload_(result) {
  const source = result || {};
  return {
    logs: Array.isArray(source.logs) ? source.logs : [],
    warnings: Array.isArray(source.warnings) ? source.warnings : [],
    lastStaffByPatient: source.lastStaffByPatient && typeof source.lastStaffByPatient === 'object' ? source.lastStaffByPatient : {},
    setupIncomplete: !!source.setupIncomplete
  };
}

function reviveTreatmentLogsCachePayload_(payload) {
  const normalized = sanitizeTreatmentLogsCachePayload_(payload);
  normalized.logs = normalized.logs.map(entry => {
    if (!entry || typeof entry !== 'object') return entry;
    const revived = Object.assign({}, entry);
    const timestamp = dashboardCoerceDate_(revived.timestamp);
    if (timestamp) {
      revived.timestamp = timestamp;
    }
    return revived;
  });
  return normalized;
}

function buildDashboardPatients_(patientInfo, sources, allowedPatientIds) {
  const patients = [];
  const basePatients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const notes = sources && sources.notes && sources.notes.notes ? sources.notes.notes : {};
  const aiReports = sources && sources.aiReports && sources.aiReports.reports ? sources.aiReports.reports : {};
  const invoices = sources && sources.invoices && sources.invoices.invoices ? sources.invoices.invoices : {};
  const responsible = sources && sources.responsible && sources.responsible.responsible ? sources.responsible.responsible : {};
  const now = dashboardCoerceDate_(sources && sources.now) || new Date();

  const seen = new Set();
  let filteredByScope = 0;
  const addPatient = (pid, payload) => {
    const patientId = dashboardNormalizePatientId_(pid);
    if (!patientId || seen.has(patientId)) return;
    if (allowedPatientIds && !allowedPatientIds.has(patientId)) {
      filteredByScope += 1;
      return;
    }
    if (!Object.prototype.hasOwnProperty.call(basePatients, patientId)) return;
    seen.add(patientId);

    const base = payload || {};
    const consentExpiryResolved = resolveConsentExpiry_(base);
    const entry = {
      patientId,
      name: base.name || base.patientName || '',
      consentExpiry: consentExpiryResolved.value == null ? '' : consentExpiryResolved.value,
      responsible: Object.prototype.hasOwnProperty.call(responsible, patientId) ? responsible[patientId] : null,
      invoiceUrl: Object.prototype.hasOwnProperty.call(invoices, patientId) ? invoices[patientId] : null,
      aiReportAt: Object.prototype.hasOwnProperty.call(aiReports, patientId) ? aiReports[patientId] : null,
      note: normalizeDashboardNote_(notes[patientId], patientId),
      statusTags: buildDashboardPatientStatusTags_(base, {
        aiReportAt: Object.prototype.hasOwnProperty.call(aiReports, patientId) ? aiReports[patientId] : null,
        note: normalizeDashboardNote_(notes[patientId], patientId),
        responsible: Object.prototype.hasOwnProperty.call(responsible, patientId) ? responsible[patientId] : null,
        now
      })
    };
    patients.push(entry);
  };

  Object.keys(basePatients).forEach(pid => addPatient(pid, basePatients[pid]));

  if (typeof dashboardLogContext_ === 'function') {
    dashboardLogContext_('buildDashboardPatients_:scope', JSON.stringify({
      applyFilter: !!allowedPatientIds,
      basePatients: Object.keys(basePatients).length,
      filteredByScope,
      resultPatients: patients.length
    }));
  }

  return patients;
}

function buildDashboardPatientStatusTags_(patient, params, maybeNow) {
  const tags = [];
  const options = params && typeof params === 'object' && !(params instanceof Date)
    ? params
    : { aiReportAt: params, now: maybeNow };
  const targetNow = dashboardCoerceDate_(options.now) || new Date();
  const aiReportAt = options.aiReportAt;
  const consentExpiryResolved = resolveConsentExpiry_(patient);
  const consentExpiryDate = parseConsentDate_(consentExpiryResolved.value);
  const raw = patient && patient.raw ? patient.raw : null;
  const consentAcquired = dashboardIsConsentAcquired_(raw);
  const daysRemaining = consentExpiryDate ? dashboardDaysBetween_(targetNow, consentExpiryDate, true) : null;

  if (shouldDebugConsent_() && consentExpiryResolved.value != null && !consentExpiryDate) {
    dashboardLogContext_('consent-date-parse-failed', JSON.stringify({
      phase: 'tag',
      patientId: patient && patient.patientId ? patient.patientId : '',
      source: consentExpiryResolved.source,
      raw: consentExpiryResolved.value
    }));
  }

  if (!consentAcquired && consentExpiryDate) {
    if (daysRemaining > 30) {
      return tags;
    }
    let label = '📄 要対応';
    let priority = 'low';
    if (daysRemaining < 0) {
      label = '⚠ 期限超過';
      priority = 'high';
    } else if (daysRemaining <= 14) {
      label = '⏳ 期限迫る';
      priority = 'medium';
    } else if (daysRemaining <= 30) {
      label = '📄 要対応';
      priority = 'low';
    }
    tags.push({ type: 'consent', label, priority });
  }

  const reportDate = dashboardParseTimestamp_(aiReportAt);
  if (!consentAcquired) {
    tags.push({ type: 'report', label: reportDate ? '作成済' : '未作成' });
  }

  return tags;
}


function dashboardDaysBetween_(from, to, futurePositive) {
  const start = dashboardCoerceDate_(from);
  const end = dashboardCoerceDate_(to);
  if (!start || !end) return Number.POSITIVE_INFINITY;
  const diff = end.getTime() - start.getTime();
  const days = Math.floor(diff / (24 * 60 * 60 * 1000));
  return futurePositive ? days : Math.abs(days);
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
  const invoiceScope = scope;

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
    invoices,
    treatmentLogs,
    payload.notes,
    invoiceScope,
    patientNameMap,
    now,
    tz,
    patientInfo
  );
  const consentRelated = buildOverviewFromConsent_(patientInfo, scope, patientNameMap, now);
  const visitSummary = buildOverviewFromTreatmentProgress_(payload.visits, now, tz);
  return {
    invoiceUnconfirmed,
    consentRelated,
    visitSummary
  };
}

function buildOverviewFromInvoiceUnconfirmed_(invoices, treatmentLogs, notes, scope, patientNameMap, now, tz, patientInfo) {
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
  let filteredCount = 0;
  (treatmentLogs || []).forEach(entry => {
    const pid = dashboardNormalizePatientId_(entry && entry.patientId);
    if (!pid) return;
    if (allowedPatientIds && !allowedPatientIds.has(pid)) {
      filteredCount += 1;
      return;
    }
    const monthKey = resolveDashboardMonthKey_(entry, tz);
    if (monthKey !== previousMonthKey) return;
    prevMonthPatientIds.add(pid);
  });
  logBillingDebug(`[billing-scope] visiblePatientIds size=${allowedPatientIds ? allowedPatientIds.size : 'all'} filteredCount=${filteredCount}`);
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
      subText: `受渡未確認（対象月: ${previousMonthKey}）`
    });
  });

  const pendingSample = items.slice(0, 10).map(item => item.patientId);
  logBillingDebug(`pendingPatients count=${items.length}`, `sample=${JSON.stringify(pendingSample)}`);

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  return { items };
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

function buildOverviewFromConsent_(patientInfo, scope, patientNameMap, now) {
  const items = [];
  const allowedPatientIds = scope ? scope.patientIds : null;
  const applyFilter = scope ? scope.applyFilter : false;
  const targetNow = dashboardCoerceDate_(now) || new Date();
  let filteredByScope = 0;

  Object.keys(patientInfo || {}).forEach(pid => {
    if (!pid) return;
    if (applyFilter && allowedPatientIds && !allowedPatientIds.has(pid)) {
      filteredByScope += 1;
      return;
    }
    const info = patientInfo[pid] || {};
    const consentExpiryResolved = resolveConsentExpiry_(info);
    const consentExpiryDate = parseConsentDate_(consentExpiryResolved.value);
    const consentAcquired = dashboardIsConsentAcquired_(info.raw);
    if (shouldDebugConsent_() && consentExpiryResolved.value != null && !consentExpiryDate) {
      dashboardLogContext_('consent-date-parse-failed', JSON.stringify({
        phase: 'overview',
        patientId: pid,
        source: consentExpiryResolved.source,
        raw: consentExpiryResolved.value
      }));
    }
    if (consentAcquired || !consentExpiryDate) return;

    const todayStart = new Date(targetNow.getFullYear(), targetNow.getMonth(), targetNow.getDate());
    const expiryStart = new Date(consentExpiryDate.getFullYear(), consentExpiryDate.getMonth(), consentExpiryDate.getDate());
    const diffDays = Math.floor((expiryStart.getTime() - todayStart.getTime()) / (24 * 60 * 60 * 1000));
    if (diffDays > 30) return;
    let label = `同意期限（残${diffDays}日）`;
    if (diffDays < 0) {
      label = `⚠ 同意期限超過（${Math.abs(diffDays)}日超過）`;
    } else if (diffDays <= 14) {
      label = `⏳ 同意期限迫る（残${diffDays}日）`;
    } else if (diffDays <= 30) {
      label = `同意期限（残${diffDays}日）`;
    }
    const name = info.name || patientNameMap[pid] || '';
    items.push({
      patientId: pid,
      name,
      subText: label
    });
  });

  items.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));
  if (typeof dashboardLogContext_ === 'function') {
    dashboardLogContext_('buildOverviewFromConsent_:scope', JSON.stringify({
      applyFilter,
      totalPatients: Object.keys(patientInfo || {}).length,
      filteredByScope,
      resultItems: items.length
    }));
  }
  return { items };
}

function resolveConsentExpiry_(patient) {
  const info = patient && typeof patient === 'object' ? patient : {};
  const raw = info.raw && typeof info.raw === 'object' ? info.raw : null;
  const candidates = [
    { source: 'info.consentExpiry', value: info.consentExpiry },
    { source: "raw['同意期限']", value: raw ? raw['同意期限'] : null },
    { source: "raw['同意有効期限']", value: raw ? raw['同意有効期限'] : null },
    { source: "raw['同意期限日']", value: raw ? raw['同意期限日'] : null }
  ];

  for (let i = 0; i < candidates.length; i++) {
    const entry = candidates[i];
    if (entry.value == null) continue;
    if (typeof entry.value === 'string' && !entry.value.trim()) continue;
    if (typeof dashboardLogContext_ === 'function') {
      dashboardLogContext_('resolveConsentExpiry_:result', JSON.stringify({
        source: entry.source,
        resolvedValue: entry.value
      }));
    }
    return { value: entry.value, source: entry.source };
  }

  if (typeof dashboardLogContext_ === 'function') {
    dashboardLogContext_('resolveConsentExpiry_:result', JSON.stringify({
      source: '',
      resolvedValue: null
    }));
  }

  return { value: null, source: '' };
}

function parseConsentDate_(value) {
  const logParseResult = result => {
    if (typeof dashboardLogContext_ !== 'function') return;
    dashboardLogContext_('parseConsentDate_:result', JSON.stringify({
      input: value,
      result: result ? result.toISOString() : null
    }));
  };

  if (value instanceof Date) {
    const parsedDate = Number.isNaN(value.getTime()) ? null : value;
    logParseResult(parsedDate);
    return parsedDate;
  }
  if (value == null) {
    logParseResult(null);
    return null;
  }

  const str = String(value).trim();
  if (!str) {
    logParseResult(null);
    return null;
  }

  const ymdHyphen = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (ymdHyphen) {
    const parsedDate = createDateFromYmd_(ymdHyphen[1], ymdHyphen[2], ymdHyphen[3]);
    logParseResult(parsedDate);
    return parsedDate;
  }

  const ymdSlash = str.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (ymdSlash) {
    const parsedDate = createDateFromYmd_(ymdSlash[1], ymdSlash[2], ymdSlash[3]);
    logParseResult(parsedDate);
    return parsedDate;
  }

  const ymdJapanese = str.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (ymdJapanese) {
    const parsedDate = createDateFromYmd_(ymdJapanese[1], ymdJapanese[2], ymdJapanese[3]);
    logParseResult(parsedDate);
    return parsedDate;
  }

  const isoPattern = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(?::\d{2}(?:\.\d{1,3})?)?(?:Z|[+-]\d{2}:\d{2})?$/;
  if (isoPattern.test(str)) {
    const timestamp = Date.parse(str);
    if (Number.isFinite(timestamp)) {
      const parsedDate = new Date(timestamp);
      logParseResult(parsedDate);
      return parsedDate;
    }
  }

  logParseResult(null);
  return null;
}

function createDateFromYmd_(yearRaw, monthRaw, dayRaw) {
  const year = Number(yearRaw);
  const month = Number(monthRaw);
  const day = Number(dayRaw);
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;

  const date = new Date(year, month - 1, day);
  if (Number.isNaN(date.getTime())) return null;
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  return date;
}

function shouldDebugConsent_() {
  if (typeof DASHBOARD_DEBUG_CONSENT !== 'undefined') return !!DASHBOARD_DEBUG_CONSENT;
  if (typeof DEBUG_MODE !== 'undefined') return !!DEBUG_MODE;
  return false;
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

function dashboardIsConsentAcquired_(raw) {
  const value = resolvePatientRawValue_(raw, ['同意書取得確認']);
  if (value === true) return true;
  if (value == null) return false;
  if (typeof value === 'boolean') return false;
  if (typeof value === 'number') return value === 1;

  const normalized = String(value).trim();
  if (!normalized) return false;

  const trueValues = new Set(['済', '取得済', '完了', 'true', 'TRUE', '1', 'yes']);
  const falseValues = new Set(['未', '未取得', 'false', 'FALSE', '0', 'no']);

  if (trueValues.has(normalized)) return true;
  if (falseValues.has(normalized)) return false;
  return false;
}

function buildOverviewFromTreatmentProgress_(visits, now, tz) {
  const targetNow = dashboardCoerceDate_(now) || new Date();
  const todayKey = dashboardFormatDate_(targetNow, tz, 'yyyy-MM-dd');
  const normalizedVisits = Array.isArray(visits) ? visits : [];
  const countByDate = {};

  normalizedVisits.forEach(visit => {
    const dateKey = visit && visit.dateKey ? String(visit.dateKey).trim() : '';
    if (!dateKey) return;
    countByDate[dateKey] = (countByDate[dateKey] || 0) + 1;
  });


  const todayCount = countByDate[todayKey] || 0;
  const latestPastDate = Object.keys(countByDate)
    .filter(dateKey => dateKey < todayKey)
    .sort((a, b) => b.localeCompare(a))[0] || '';
  const latestDayCount = todayCount > 0 ? todayCount : (latestPastDate ? countByDate[latestPastDate] || 0 : 0);

  return {
    todayCount,
    latestDayCount
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
