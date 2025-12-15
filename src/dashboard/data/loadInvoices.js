/**
 * 当月の請求書PDFリンクを患者IDと紐付けて返す。
 * @param {Object} [options]
 * @param {Object} [options.patientInfo]
 * @param {Object} [options.nameToId]
 * @param {Object} [options.rootFolder]
 * @param {Date} [options.now]
 * @return {{invoices: Object<string, string|null>, warnings: string[]}}
 */
function loadInvoices(options) {
  const opts = options || {};
  const fetchFn = () => loadInvoicesUncached_(opts);
  if (typeof dashboardCacheFetch_ === 'function') {
    return dashboardCacheFetch_(dashboardCacheKey_('invoices:v1'), fetchFn, DASHBOARD_CACHE_TTL_SECONDS, opts);
  }
  return fetchFn();
}

function loadInvoicesUncached_(options) {
  const opts = options || {};
  const patientInfo = opts.patientInfo || (typeof loadPatientInfo === 'function' ? loadPatientInfo() : null);
  const patients = patientInfo && patientInfo.patients ? patientInfo.patients : {};
  const nameToId = opts.nameToId || (patientInfo && patientInfo.nameToId) || {};
  const warnings = patientInfo && Array.isArray(patientInfo.warnings) ? [].concat(patientInfo.warnings) : [];
  const invoices = {};
  const latestMeta = {};

  Object.keys(patients).forEach(pid => {
    const normalized = dashboardNormalizePatientId_(pid);
    if (normalized) {
      invoices[normalized] = null;
    }
  });

  const root = opts.rootFolder || dashboardGetInvoiceRootFolder_();
  if (!root || typeof root.getFolders !== 'function') {
    warnings.push('請求書フォルダが取得できませんでした');
    dashboardWarn_('[loadInvoices] invoice root folder not found');
    return { invoices, warnings };
  }

  const tz = dashboardResolveTimeZone_();
  const now = dashboardCoerceDate_(opts.now) || new Date();
  const targetYmDigits = dashboardFormatDate_(now, tz, 'yyyyMM');
  const targetYmHyphen = dashboardFormatDate_(now, tz, 'yyyy-MM');
  const targetYmKanji = dashboardFormatDate_(now, tz, 'yyyy年MM月');

  const folders = root.getFolders();
  while (folders && typeof folders.hasNext === 'function' && folders.hasNext()) {
    const folder = folders.next();
    const name = folder && typeof folder.getName === 'function' ? folder.getName() : '';
    if (!isTargetInvoiceFolder_(name, targetYmDigits, targetYmKanji)) continue;

    const files = folder.getFiles && folder.getFiles();
    while (files && typeof files.hasNext === 'function' && files.hasNext()) {
      const file = files.next();
      const fileName = file && typeof file.getName === 'function' ? file.getName() : '';
      const parsed = parseInvoiceFileName_(fileName, targetYmDigits, targetYmHyphen);
      if (!parsed) continue;

      const pid = dashboardResolvePatientIdFromName_(parsed.patientName, nameToId);
      if (!pid) {
        warnings.push(`患者名をIDに紐付けできません: ${parsed.patientName}`);
        continue;
      }

      const updated = file && typeof file.getLastUpdated === 'function' ? file.getLastUpdated() : null;
      const updatedDate = dashboardCoerceDate_(updated);
      const updatedTs = updatedDate ? updatedDate.getTime() : 0;
      const current = latestMeta[pid];
      if (current && current.updatedTs >= updatedTs) continue;

      const url = file && typeof file.getUrl === 'function' ? file.getUrl() : '';
      latestMeta[pid] = { updatedTs, url };
      invoices[pid] = url || null;
    }
  }

  return { invoices, warnings };
}

function isTargetInvoiceFolder_(name, targetYmDigits, targetYmKanji) {
  const trimmed = String(name == null ? '' : name).trim();
  if (!trimmed) return false;
  const digitPrefix = targetYmDigits ? targetYmDigits + '請求書_' : '';
  const kanjiPrefix = targetYmKanji ? targetYmKanji + '請求書_' : '';
  return (digitPrefix && trimmed.indexOf(digitPrefix) === 0)
    || (kanjiPrefix && trimmed.indexOf(kanjiPrefix) === 0);
}

function parseInvoiceFileName_(fileName, targetYmDigits, targetYmHyphen) {
  const trimmed = String(fileName == null ? '' : fileName).trim();
  const match = trimmed.match(/^(.+?)_([0-9]{4}-[0-9]{2}|[0-9]{6,8})_請求書\.pdf$/i);
  if (!match) return null;
  const rawMonth = match[2];
  const matchesDigits = targetYmDigits && rawMonth.replace(/[^0-9]/g, '').slice(0, 6) === targetYmDigits;
  const matchesHyphen = targetYmHyphen && rawMonth === targetYmHyphen;
  if (!matchesDigits && !matchesHyphen) return null;
  return { patientName: match[1] };
}
